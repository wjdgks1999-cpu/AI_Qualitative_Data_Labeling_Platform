"""합성 문서의 군집 관계와 전처리 Excel 연결을 검증한다."""

import numpy as np
import pandas as pd
import pytest

from exploration.cluster_texts import ClusterConfig, cluster_documents, main, make_cluster_figure, project_clusters
from preprocessing.preprocess_text import preprocess_responses, save_preprocessed_data


def identity(text):
    """합성 명사 문장을 외부 처리 없이 그대로 전달한다."""
    return text


def test_similar_documents_cluster_together_and_noise_is_counted():
    """유사도 대신 거리로 묶이며 서로 다른 문서와 빈 문서는 구분된다."""
    source = pd.DataFrame({"tokenized_text": ["사과 과일", "사과 과일", "자동차 도로", "자동차 도로", None]})
    before = source.copy(deep=True)
    result = cluster_documents(source, ClusterConfig(eps_values=(0.1,), min_samples_values=(2,)))
    labels = result.data["cluster"].tolist()
    assert labels[0] == labels[1] != -1
    assert labels[2] == labels[3] != -1
    assert labels[0] != labels[2]
    assert labels[4] == -1
    assert result.num_clusters == 2
    assert result.trials["num_noise"].tolist() == [1]
    pd.testing.assert_frame_equal(source, before)


def test_no_noise_keeps_full_cluster_count_and_first_tie():
    """잡음이 없으면 군집 수를 차감하지 않으며 동률에서 첫 조합을 선택한다."""
    source = pd.DataFrame({"tokenized_text": ["사과 과일", "사과 과일", "자동차 도로", "자동차 도로"]})
    result = cluster_documents(source, ClusterConfig(eps_values=(0.1, 0.5), min_samples_values=(2,)))
    assert result.num_clusters == 2
    assert result.best_eps == 0.1
    assert result.trials["num_noise"].tolist() == [0, 0]


def test_all_noise_still_has_labels_and_selected_parameters():
    """모든 문서가 잡음이어도 None 대신 -1 라벨과 선택 조합을 돌려준다."""
    source = pd.DataFrame({"tokenized_text": ["사과 과일", "자동차 도로"]})
    result = cluster_documents(source, ClusterConfig(eps_values=(0.1,), min_samples_values=(3, 5)))
    assert result.num_clusters == 0
    assert result.data["cluster"].tolist() == [-1, -1]
    assert result.best_min_samples == 3


@pytest.mark.parametrize("source", [pd.DataFrame(), pd.DataFrame({"other": ["합성"]}), pd.DataFrame({"tokenized_text": ["", None]}), pd.DataFrame({"tokenized_text": ["가", "나"]}), pd.DataFrame({"tokenized_text": [1]})])
def test_unusable_text_inputs_fail_clearly(source):
    """빈 자료·잘못된 열·사용할 어휘가 없는 문서·숫자 입력을 거부한다."""
    with pytest.raises(ValueError):
        cluster_documents(source)


@pytest.mark.parametrize("config", [ClusterConfig(eps_values=()), ClusterConfig(eps_values=(-1,)), ClusterConfig(eps_values=(float("nan"),)), ClusterConfig(min_samples_values=()), ClusterConfig(min_samples_values=(0,))])
def test_invalid_grid_is_rejected(config):
    """유효하지 않은 파라미터 조합은 탐색 전에 거부한다."""
    with pytest.raises(ValueError):
        cluster_documents(pd.DataFrame({"tokenized_text": ["합성 문서"]}), config)


def test_preprocessing_excel_can_be_read_by_clustering(tmp_path, capsys):
    """대체 전처리 결과를 실제 Excel로 저장해 군집 명령에 전달한다."""
    data = preprocess_responses(["사과 과일!", "사과 과일?", "자동차 도로!", "자동차 도로?"], spacing=identity, spellcheck=identity, tokenize=identity)
    path = save_preprocessed_data(data, tmp_path / "preprocessing_data.xlsx")
    before = path.read_bytes()
    assert main([str(path), "--no-plot"]) == 0
    assert path.read_bytes() == before
    output = capsys.readouterr().out
    assert "군집 수=2" in output
    assert "사과" not in output


def test_tsne_sample_limit_and_actual_figure():
    """표본 수 조건을 확인하고 실제 t-SNE와 비대화형 그림 생성을 검증한다."""
    source = pd.DataFrame({"tokenized_text": ["사과 과일", "사과 맛", "자동차 도로", "자동차 운전", "음악 노래", "음악 연주"]})
    result = cluster_documents(source)
    with pytest.raises(ValueError, match="perplexity"):
        project_clusters(result)
    coordinates = project_clusters(result, perplexity=2)
    assert coordinates.shape == (6, 2)
    assert np.isfinite(coordinates).all()
    figure = make_cluster_figure(result, coordinates)
    figure.canvas.draw()
    np.testing.assert_array_equal(figure.axes[0].collections[0].get_array(), result.data["cluster"].to_numpy())
    import matplotlib.pyplot as plt

    plt.close(figure)
