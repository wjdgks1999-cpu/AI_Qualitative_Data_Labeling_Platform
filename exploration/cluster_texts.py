"""전처리된 응답을 TF-IDF와 DBSCAN으로 군집 탐색한다."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.cluster import DBSCAN
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_distances, cosine_similarity


@dataclass(frozen=True)
class ClusterConfig:
    """기존 TF-IDF 설정과 DBSCAN 탐색 범위를 보관한다."""

    min_df: float = 0.0
    max_df: float = 1.0
    max_features: int = 50
    ngram_range: tuple[int, int] = (1, 2)
    eps_values: tuple[float, ...] = (0.1, 0.5, 1.0)
    min_samples_values: tuple[int, ...] = (2, 3, 5)


@dataclass
class ClusterResult:
    """군집을 붙인 응답과 조합별 결과, 시각화용 TF-IDF 행렬을 담는다."""

    data: pd.DataFrame
    trials: pd.DataFrame
    features: object
    best_eps: float
    best_min_samples: int
    num_clusters: int


def cluster_documents(data: pd.DataFrame, config: ClusterConfig | None = None) -> ClusterResult:
    """코사인 거리로 조합을 비교하고 잡음을 제외한 군집 수가 가장 큰 결과를 반환한다.

    동률에서는 탐색 순서상 첫 조합을 유지한다. 전부 잡음인 경우에도
    첫 조합과 -1 라벨을 반환한다. 원본 DataFrame은 변경하지 않는다.
    """
    settings = config or ClusterConfig()
    if "tokenized_text" not in data.columns or data.empty:
        raise ValueError("비어 있지 않은 데이터와 tokenized_text 열이 필요합니다.")
    if not settings.eps_values or not settings.min_samples_values:
        raise ValueError("DBSCAN 탐색 범위가 비어 있습니다.")
    if any(not np.isfinite(value) or value <= 0 for value in settings.eps_values):
        raise ValueError("eps는 0보다 큰 유한수여야 합니다.")
    if any(not isinstance(value, int) or isinstance(value, bool) or value < 1 for value in settings.min_samples_values):
        raise ValueError("min_samples는 양의 정수여야 합니다.")
    result_data = data.copy()
    result_data["tokenized_text"] = result_data["tokenized_text"].fillna("")
    if not result_data["tokenized_text"].map(lambda value: isinstance(value, str)).all():
        raise ValueError("tokenized_text는 문자열 또는 결측값이어야 합니다.")
    vectorizer = TfidfVectorizer(min_df=settings.min_df, max_df=settings.max_df, max_features=settings.max_features, ngram_range=settings.ngram_range)
    try:
        features = vectorizer.fit_transform(result_data["tokenized_text"])
    except ValueError as error:
        if "empty vocabulary" in str(error):
            raise ValueError("TF-IDF로 사용할 단어가 없습니다. 빈 응답과 한 글자 토큰을 확인하세요.") from None
        raise
    distances = cosine_distances(features)
    trials = []
    best = None
    for eps in settings.eps_values:
        for min_samples in settings.min_samples_values:
            labels = DBSCAN(eps=eps, min_samples=min_samples, metric="precomputed").fit_predict(distances)
            num_clusters = len(set(labels) - {-1})
            trials.append({"eps": eps, "min_samples": min_samples, "num_clusters": num_clusters, "num_noise": int(np.count_nonzero(labels == -1))})
            if best is None or num_clusters > best[0]:
                best = (num_clusters, eps, min_samples, labels)
    num_clusters, eps, min_samples, labels = best
    result_data["cluster"] = labels
    return ClusterResult(result_data, pd.DataFrame(trials), features, eps, min_samples, num_clusters)


def project_clusters(result: ClusterResult, *, perplexity: float = 30.0) -> np.ndarray:
    """기존처럼 코사인 유사도 행벡터를 t-SNE의 특징 입력으로 사용한다.

    DBSCAN의 거리 수정과 시각화의 의미 변경을 구분하기 위해 시각화
    입력 표현은 유지한다. 표본 수 미만의 perplexity만 허용한다.
    """
    if not np.isfinite(perplexity) or not 0 < perplexity < len(result.data):
        raise ValueError("perplexity는 0보다 크고 응답 수보다 작아야 합니다.")
    from sklearn.manifold import TSNE

    similarities = cosine_similarity(result.features)
    return TSNE(n_components=2, random_state=42, perplexity=perplexity).fit_transform(similarities)


def make_cluster_figure(result: ClusterResult, coordinates: np.ndarray):
    """계산된 좌표와 군집 라벨로 그림을 만들며 저장이나 화면 표시는 하지 않는다."""
    import matplotlib.pyplot as plt

    if coordinates.shape != (len(result.data), 2):
        raise ValueError("시각화 좌표는 응답 수 × 2 형태여야 합니다.")
    figure, axes = plt.subplots(figsize=(8, 6))
    scatter = axes.scatter(coordinates[:, 0], coordinates[:, 1], c=result.data["cluster"], cmap="viridis", alpha=0.7)
    axes.set_title("t-SNE Visualization of Clusters")
    axes.legend(*scatter.legend_elements(), title="Clusters")
    return figure


def main(argv: list[str] | None = None) -> int:
    """Excel에서 군집을 탐색하고 응답 원문 없이 조합별 요약과 선택적 그래프를 표시한다."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, nargs="?", default=Path("preprocessing_data.xlsx"))
    parser.add_argument("--eps", type=float, nargs="+", default=list(ClusterConfig.eps_values))
    parser.add_argument("--min-samples", type=int, nargs="+", default=list(ClusterConfig.min_samples_values))
    parser.add_argument("--max-features", type=int, default=50)
    parser.add_argument("--perplexity", type=float, default=30.0)
    parser.add_argument("--no-plot", action="store_true", help="t-SNE 계산과 그래프 표시 생략")
    args = parser.parse_args(argv)
    try:
        data = pd.read_excel(args.input)
        size_mib = len(data) ** 2 * 8 / 1024 ** 2
        print(f"응답 수: {len(data)}, 밀집 거리 행렬 하나의 예상 크기: {size_mib:.1f} MiB (전체 메모리 사용량은 더 큼)")
        settings = ClusterConfig(max_features=args.max_features, eps_values=tuple(args.eps), min_samples_values=tuple(args.min_samples))
        result = cluster_documents(data, settings)
        print(result.trials.to_string(index=False))
        print(f"선택 조합: eps={result.best_eps}, min_samples={result.best_min_samples}, 군집 수={result.num_clusters}")
        if not args.no_plot:
            coordinates = project_clusters(result, perplexity=args.perplexity)
            make_cluster_figure(result, coordinates)
            import matplotlib.pyplot as plt

            plt.show()
    except Exception as error:
        print(f"군집 탐색 실패: {type(error).__name__}. 입력 열·문서 내용·파라미터·표본 수를 확인하세요.")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
