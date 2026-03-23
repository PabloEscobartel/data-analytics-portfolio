# 🎵 ML: Предсказание успеха музыкальных треков

## Описание
Построение и сравнение ML-моделей для предсказания успеха музыкальных треков на основе аудио-характеристик. Полный цикл: EDA → feature engineering → обучение → оценка → кластеризация.

## Что сделано

### Supervised Learning
- **KNN** — подбор оптимального k (перебор 1–25), визуализация train/test accuracy
- **Линейная регрессия** — оценка R², MSE, кросс-валидация (6-fold)
- **Логистическая регрессия** — ROC-кривая, AUC, предсказание вероятностей
- **Random Forest** — сравнение accuracy с остальными моделями

### Unsupervised Learning
- **KMeans** — подбор числа кластеров через метод локтя (elbow method)
- Анализ кластерной структуры данных

### Оценка качества
- Кросс-валидация (KFold, k=6) для всех моделей
- Сравнительная таблица метрик: R², accuracy, MSE, AUC
- ROC-кривая для логистической регрессии

## Стек
- **Python:** pandas, numpy, matplotlib, seaborn
- **ML:** scikit-learn (KNeighborsClassifier, LinearRegression, LogisticRegression, RandomForestClassifier, KMeans, StandardScaler, cross_val_score, roc_curve, roc_auc_score)
- **Среда:** Google Colab
