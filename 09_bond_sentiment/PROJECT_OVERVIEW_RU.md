# Описание проекта и навигация по папкам

Этот архив содержит материалы проекта о связи публичного сентимента до букбилдинга с результатами размещения корпоративных облигаций. В проект входят итоговый отчет, презентация, подготовленные данные, код для NLP-классификации сообщений, скрипты эконометрического анализа и итоговые файлы с результатами.

## Что открыть в первую очередь

Для чтения результатов без запуска кода:

- `report_final.docx` - итоговый текст работы.
- `Защита_сем_наставника.pptx` - презентация.
- `final_results_v3_current/` - основные итоговые результаты для базового окна агрегации сентимента 14 дней.
- `final_results_lookback7/` - проверка робастности для окна агрегации сентимента 7 дней.
- `final_results_v3_current/results_workbooks_roadmap.md` - карта Excel-файлов с результатами.
- `final_results_lookback7/results_workbooks_roadmap_lookback7.md` - карта Excel-файлов для 7-дневного окна.

## Основная логика проекта

Единица наблюдения - выпуск корпоративной облигации. Для каждого выпуска собирается панель с параметрами размещения, кредитным рейтингом, сроком, ставкой ОФЗ, историей предыдущих размещений эмитента и другими контролями.

Затем к выпуску агрегируются сообщения о соответствующем эмитенте за окно до открытия или закрытия книги заявок. В основной спецификации используется окно 14 дней до события, в робастности - 7 дней. Сентимент агрегируется только по релевантным сообщениям, прошедшим NLP-классификацию.

Основные зависимые переменные:

- утаптывание купона/спреда: `coupon_reduction_bp`;
- факт утаптывания купона: `coupon_reduced_dummy`;
- увеличение объема размещения относительно тизера: `placement_vol_book = placement_volume / teaser_volume`;
- лог объема относительно тизера: `log_placement_vol_book`;
- dummy увеличения объема: `upsize_dummy = 1[placement_vol_book > 1.01]`;
- совместный успех: утаптывание купона и увеличение объема одновременно.

Основные переменные интереса:

- `si` - индекс сентимента;
- `log_buzz` - логарифм числа релевантных сообщений;
- split-компоненты позитивного и негативного сентимента в дополнительных спецификациях.

## Где лежат входные данные

Основные входы для финального аналитического пайплайна находятся здесь:

```text
analysis_pipeline_final/data_inputs/
```

Ключевые файлы:

- `panel_final_v3.xlsx` - финальная панель выпусков облигаций.
- `classified.csv` - сообщения, прошедшие NLP-классификацию релевантности и сентимента.

Внутри `panel_final_v3.xlsx` есть лист:

- `panel` - данные по выпускам;
- `description` - описание части переменных панели.

## Где искать итоговые результаты

### Основные результаты, окно 14 дней

```text
final_results_v3_current/
```

Главные файлы:

- `regression_results_upsize_all_windows.xlsx` - основные регрессии на полной выборке.
- `regression_results_upsize_all_windows_fixed.xlsx` - регрессии отдельно для выпусков с фиксированным купоном.
- `regression_results_upsize_all_windows_floater.xlsx` - регрессии отдельно для флоутеров.
- `diagnostics_results_upsize_all_windows.xlsx` - диагностические тесты для полной выборки.
- `diagnostics_results_upsize_all_windows_fixed.xlsx` - диагностика для фиксированных купонов.
- `diagnostics_results_upsize_all_windows_floater.xlsx` - диагностика для флоутеров.
- `descriptive_statistics_all_windows.xlsx` - описательные статистики.
- `distribution_plots_all_windows.pdf` - графики распределений переменных.
- `prediction_results_book_open_all_thresholds.xlsx` - предиктивные модели только на данных до открытия книги.
- `gjrm_biprobit_bootstrap_coefficients_all.xlsx` - bivariate probit/GJRM для полной выборки.
- `gjrm_biprobit_bootstrap_coefficients_fixed.xlsx` - GJRM для фиксированных купонов.
- `gjrm_biprobit_bootstrap_coefficients_floater.xlsx` - GJRM для флоутеров.
- `regression_results_high_spread_open_5.xlsx` - проверка для высокодоходного сегмента.
- `results_workbooks_roadmap.md` - описание листов внутри итоговых Excel-файлов.

### Робастность, окно 7 дней

```text
final_results_lookback7/
```

Структура аналогична основной папке, но результаты построены при агрегации сентимента за 7 дней до открытия/закрытия книги.

Главные файлы:

- `regression_results_upsize_lookback7_all_windows.xlsx`;
- `regression_results_upsize_lookback7_all_windows_fixed.xlsx`;
- `regression_results_upsize_lookback7_all_windows_floater.xlsx`;
- `diagnostics_results_upsize_lookback7_all_windows.xlsx`;
- `descriptive_statistics_lookback7_all_windows.xlsx`;
- `distribution_plots_lookback7_all_windows.pdf`;
- `prediction_results_book_open_lookback7_all_thresholds.xlsx`;
- `gjrm_biprobit_bootstrap_coefficients_lookback7_all.xlsx`;
- `gjrm_biprobit_bootstrap_coefficients_lookback7_fixed.xlsx`;
- `gjrm_biprobit_bootstrap_coefficients_lookback7_floater.xlsx`;
- `regression_results_high_spread_lookback7_open_5.xlsx`;
- `results_workbooks_roadmap_lookback7.md`.

### Другие папки с результатами

- `final_results/` - более ранний набор результатов, сохранен для трассировки.
- `final_results_v3_open_predictions/` - промежуточный набор после перехода prediction-блока на данные до открытия книги.

Для цитирования в работе лучше использовать `final_results_v3_current/` и `final_results_lookback7/`.

## Где лежат скрипты регрессий и диагностики

Финальный аналитический код находится здесь:

```text
analysis_pipeline_final/scripts/
```

Основные скрипты:

- `build_pipeline.py` - создает regression-ready датасеты из `panel_final_v3.xlsx` и `classified.csv`.
- `run_regressions_upsize.py` - строит один набор регрессий.
- `run_regressions_upsize_batch.py` - запускает регрессии по всем окнам, порогам сообщений и подвыборкам.
- `run_diagnostics_upsize.py` - диагностические тесты для одного датасета.
- `run_diagnostics_upsize_batch.py` - диагностические тесты по всем спецификациям.
- `run_descriptive_statistics.py` - описательные статистики и графики распределений.
- `run_prediction.py` - ML/prediction-блок для одного датасета.
- `run_prediction_batch.py` - batch-запуск prediction-моделей.
- `run_regressions_high_spread.py` - проверка для высокодоходного сегмента.
- `run_biprobit_gjrm_cluster_bootstrap_v4.R` - bivariate probit/GJRM с кластерным бутстрапом.
- `combine_gjrm_bootstrap_se.py` - объединение результатов бутстрапа GJRM.

## Как запустить финальный пайплайн

Готовые команды находятся в:

```text
analysis_pipeline_final/runbooks/
```

Основное окно 14 дней:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_14d.sh
```

Робастность 7 дней:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_7d.sh
```

Сбор результатов в итоговые папки:

```bash
bash analysis_pipeline_final/runbooks/collect_final_results.sh 14d
bash analysis_pipeline_final/runbooks/collect_final_results.sh 7d
```

## Где лежит NLP-часть

NLP-блок находится здесь:

```text
train NLP/
```

В этой папке находятся:

- скрипты мэтчинга сообщений к эмитентам и выпускам;
- данные для разметки;
- файлы с классифицированными сообщениями;
- скрипты обучения моделей релевантности и сентимента;
- обученные модели, если они включены в передаваемый архив.

Ключевые файлы:

- `bond_entity_matcher_v10.py` - основной мэтчинг сообщений к эмитентам/выпускам.
- `prepare_primary_sentiment_universe_close_time.py` - подготовка universe сообщений для классификации.
- `sample_for_labeling.py` - формирование выборки для разметки.
- `classify_samples_gemini.py` - классификация/разметка с использованием Gemini.
- `train_bond_classifier_final.py` - обучение классификаторов.
- `classify_sentiment_gemini_genai_resume_from_csv.py` - применение классификации к сообщениям.
- `classified.csv` - итоговый файл с классифицированными сообщениями, который подается в аналитический пайплайн.
- `panel_final_v3.xlsx` - копия финальной панели для NLP-блока.

Дополнительное описание NLP-папки:

```text
train NLP/README.md
train NLP/DATA_BUNDLE.md
```

## Где лежит подготовка панели

Скрипты, которые использовались до финального анализа для обогащения панели, лежат в:

```text
panel_enrichment_pipeline/
cbonds_gemini_extraction_pipeline/
raw_social_data_pipeline/
```

Назначение папок:

- `panel_enrichment_pipeline/` - рейтинги, организаторы, параметры выпусков, дополнительные поля панели.
- `cbonds_gemini_extraction_pipeline/` - извлечение ориентиров купона/спреда, объемов размещения и данных букбилдинга из Cbonds/Gemini.
- `raw_social_data_pipeline/` - парсеры Telegram и Smart-Lab.

Эти папки нужны для воспроизведения upstream-сбора данных. Для проверки основных результатов достаточно `analysis_pipeline_final/`, `train NLP/` и папок `final_results*`.

## Что не является актуальной финальной спецификацией

Папка:

```text
legacy_or_superseded_scripts/
```

содержит старые или замененные версии скриптов. Она оставлена для прозрачности, но не должна использоваться как основной источник результатов.

Также отдельные старые файлы и папки результатов могут быть сохранены для истории. Если есть сомнения, ориентироваться нужно на:

```text
analysis_pipeline_final/
final_results_v3_current/
final_results_lookback7/
train NLP/
```

## Краткая карта проекта

```text
bond_sentiment_project_final/
├── report_final.docx
├── Защита_сем_наставника.pptx
├── PROJECT_OVERVIEW_RU.md
├── analysis_pipeline_final/
│   ├── data_inputs/
│   ├── scripts/
│   ├── runbooks/
│   └── docs/
├── train NLP/
├── final_results_v3_current/
├── final_results_lookback7/
├── panel_enrichment_pipeline/
├── cbonds_gemini_extraction_pipeline/
├── raw_social_data_pipeline/
└── legacy_or_superseded_scripts/
```

## Минимальный путь для проверки результатов

1. Открыть `report_final.docx`.
2. Для таблиц и расчетов открыть `final_results_v3_current/results_workbooks_roadmap.md`.
3. По roadmap открыть нужный Excel-файл из `final_results_v3_current/`.
4. Для робастности по 7-дневному окну открыть `final_results_lookback7/results_workbooks_roadmap_lookback7.md`.
5. Если нужно проверить код, смотреть `analysis_pipeline_final/scripts/`.
6. Если нужно проверить NLP-подготовку сообщений, смотреть `train NLP/`.
