library(dplyr)
library(tidyr)
library(ggplot2)
library(ggpubr)
library(car)
library(lmtest)
library(psych)
library(corrplot)
library(plm)
library(broom)
library(erer)
library(MASS)
library(AER)
library(memisc)
library(readxl)
library(stargazer)

data_meet <- read_excel("data/diplom.xlsx", sheet = "all_meets")
data_mess <- read_excel("data/diplom.xlsx", sheet = "all_mess")
data_meet_before2014 <- read_excel("data/diplom.xlsx", sheet = "before2014_meet")
data_meet_after2014 <- read_excel("data/diplom.xlsx", sheet = "after2014_meet")
data_mess_before2014 <- read_excel("data/diplom.xlsx", sheet = "before2014_mess")
data_mess_after2014 <- read_excel("data/diplom.xlsx", sheet = "after2014_mess")

# Оглавление
# Выбросы
# Гистограммы
# Матрицы корреляций
# Модели
#   все собрания: 1) решения 07; 2) рассылка 07; 3) решения 01; 4) рассылка 01; 5) решения -2+2; 6) рассылка -2+2
#   годовые собрания
#   Внеочередные собрания
# Проверка условий Гаусса-маркова
#   1) Проверка спецификации (квадраты и кубы); 2) Равенство мат ожиданий остатков нулю; 3) Проверка гомоскедастичности; 4) Проверка автокорреляции; 5) Проверка мультиколлинеарности; 6) Проверка эндогенности
# Модели до 14 и после 14 
# 
# 
# Финальные модели

summary(data_meet)
summary(data_mess)
summary(data_meet_before2014)
summary(data_mess_before2014)
summary(data_meet_after2014)
summary(data_mess_after2014)

# !!!Обработка выбросов!!!

# data_meet и data_mess

Q10_CAR_S_07_meet <- quantile(data_meet$CAR_S_07_meet, 0.01)
Q90_CAR_S_07_meet <- quantile(data_meet$CAR_S_07_meet, 0.99)
IQR_value_CAR_S_07_meet <- IQR(data_meet$CAR_S_07_meet)
lower_bound_CAR_S_07_meet <- Q10_CAR_S_07_meet - 1.5*IQR_value_CAR_S_07_meet
upper_bound_CAR_S_07_meet <- Q90_CAR_S_07_meet + 1.5*IQR_value_CAR_S_07_meet
data_meet$CAR_S_07_meet <- ifelse(data_meet$CAR_S_07_meet < lower_bound_CAR_S_07_meet, Q10_CAR_S_07_meet, ifelse(data_meet$CAR_S_07_meet > upper_bound_CAR_S_07_meet, Q90_CAR_S_07_meet, data_meet$CAR_S_07_meet))
Q10_CAR_S_07_mess <- quantile(data_mess$CAR_S_07_mess, 0.01)
Q90_CAR_S_07_mess <- quantile(data_mess$CAR_S_07_mess, 0.99)
IQR_value_CAR_S_07_mess <- IQR(data_mess$CAR_S_07_mess)
lower_bound_CAR_S_07_mess <- Q10_CAR_S_07_mess - 1.5*IQR_value_CAR_S_07_mess
upper_bound_CAR_S_07_mess <- Q90_CAR_S_07_mess + 1.5*IQR_value_CAR_S_07_mess
data_mess$CAR_S_07_mess <- ifelse(data_mess$CAR_S_07_mess < lower_bound_CAR_S_07_mess, Q10_CAR_S_07_mess, ifelse(data_mess$CAR_S_07_mess > upper_bound_CAR_S_07_mess, Q90_CAR_S_07_mess, data_mess$CAR_S_07_mess))

Q10_CAR_S_01_meet <- quantile(data_meet$CAR_S_01_meet, 0.01)
Q90_CAR_S_01_meet <- quantile(data_meet$CAR_S_01_meet, 0.99)
IQR_value_CAR_S_01_meet <- IQR(data_meet$CAR_S_01_meet)
lower_bound_CAR_S_01_meet <- Q10_CAR_S_01_meet - 1.5*IQR_value_CAR_S_01_meet
upper_bound_CAR_S_01_meet <- Q90_CAR_S_01_meet + 1.5*IQR_value_CAR_S_01_meet
data_meet$CAR_S_01_meet <- ifelse(data_meet$CAR_S_01_meet < lower_bound_CAR_S_01_meet, Q10_CAR_S_01_meet, ifelse(data_meet$CAR_S_01_meet > upper_bound_CAR_S_01_meet, Q90_CAR_S_01_meet, data_meet$CAR_S_01_meet))
Q10_CAR_S_01_mess <- quantile(data_mess$CAR_S_01_mess, 0.01)
Q90_CAR_S_01_mess <- quantile(data_mess$CAR_S_01_mess, 0.99)
IQR_value_CAR_S_01_mess <- IQR(data_mess$CAR_S_01_mess)
lower_bound_CAR_S_01_mess <- Q10_CAR_S_01_mess - 1.5*IQR_value_CAR_S_01_mess
upper_bound_CAR_S_01_mess <- Q90_CAR_S_01_mess + 1.5*IQR_value_CAR_S_01_mess
data_mess$CAR_S_01_mess <- ifelse(data_mess$CAR_S_01_mess < lower_bound_CAR_S_01_mess, Q10_CAR_S_01_mess, ifelse(data_mess$CAR_S_01_mess > upper_bound_CAR_S_01_mess, Q90_CAR_S_01_mess, data_mess$CAR_S_01_mess))

Q10_CAR_S_22_meet <- quantile(data_meet$CAR_S_22_meet, 0.01)
Q90_CAR_S_22_meet <- quantile(data_meet$CAR_S_22_meet, 0.99)
IQR_value_CAR_S_22_meet <- IQR(data_meet$CAR_S_22_meet)
lower_bound_CAR_S_22_meet <- Q10_CAR_S_22_meet - 1.5*IQR_value_CAR_S_22_meet
upper_bound_CAR_S_22_meet <- Q90_CAR_S_22_meet + 1.5*IQR_value_CAR_S_22_meet
data_meet$CAR_S_22_meet <- ifelse(data_meet$CAR_S_22_meet < lower_bound_CAR_S_22_meet, Q10_CAR_S_22_meet, ifelse(data_meet$CAR_S_22_meet > upper_bound_CAR_S_22_meet, Q90_CAR_S_22_meet, data_meet$CAR_S_22_meet))
Q10_CAR_S_22_mess <- quantile(data_mess$CAR_S_22_mess, 0.01)
Q90_CAR_S_22_mess <- quantile(data_mess$CAR_S_22_mess, 0.99)
IQR_value_CAR_S_22_mess <- IQR(data_mess$CAR_S_22_mess)
lower_bound_CAR_S_22_mess <- Q10_CAR_S_22_mess - 1.5*IQR_value_CAR_S_22_mess
upper_bound_CAR_S_22_mess <- Q90_CAR_S_22_mess + 1.5*IQR_value_CAR_S_22_mess
data_mess$CAR_S_22_mess <- ifelse(data_mess$CAR_S_22_mess < lower_bound_CAR_S_22_mess, Q10_CAR_S_22_mess, ifelse(data_mess$CAR_S_22_mess > upper_bound_CAR_S_22_mess, Q90_CAR_S_22_mess, data_mess$CAR_S_22_mess))

ggplot(data_meet, aes(x = CAR_S_07_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet simple CAR 0 +7")
ggplot(data_mess, aes(x = CAR_S_07_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess simple CAR 0 +7")
ggplot(data_meet, aes(x = CAR_S_01_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet simple CAR 0 +1")
ggplot(data_mess, aes(x = CAR_S_01_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess simple CAR 0 +1")
ggplot(data_meet, aes(x = CAR_S_22_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet simple CAR -2 +2")
ggplot(data_mess, aes(x = CAR_S_22_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess simple CAR -2 +2")


Q10_CAR_L_07_meet <- quantile(data_meet$CAR_L_07_meet, 0.01)
Q90_CAR_L_07_meet <- quantile(data_meet$CAR_L_07_meet, 0.99)
IQR_value_CAR_L_07_meet <- IQR(data_meet$CAR_L_07_meet)
lower_bound_CAR_L_07_meet <- Q10_CAR_L_07_meet - 1.5*IQR_value_CAR_L_07_meet
upper_bound_CAR_L_07_meet <- Q90_CAR_L_07_meet + 1.5*IQR_value_CAR_L_07_meet
data_meet$CAR_L_07_meet <- ifelse(data_meet$CAR_L_07_meet < lower_bound_CAR_L_07_meet, Q10_CAR_L_07_meet, ifelse(data_meet$CAR_L_07_meet > upper_bound_CAR_L_07_meet, Q90_CAR_L_07_meet, data_meet$CAR_L_07_meet))
Q10_CAR_L_07_mess <- quantile(data_mess$CAR_L_07_mess, 0.01)
Q90_CAR_L_07_mess <- quantile(data_mess$CAR_L_07_mess, 0.99)
IQR_value_CAR_L_07_mess <- IQR(data_mess$CAR_L_07_mess)
lower_bound_CAR_L_07_mess <- Q10_CAR_L_07_mess - 1.5*IQR_value_CAR_L_07_mess
upper_bound_CAR_L_07_mess <- Q90_CAR_L_07_mess + 1.5*IQR_value_CAR_L_07_mess
data_mess$CAR_L_07_mess <- ifelse(data_mess$CAR_L_07_mess < lower_bound_CAR_L_07_mess, Q10_CAR_L_07_mess, ifelse(data_mess$CAR_L_07_mess > upper_bound_CAR_L_07_mess, Q90_CAR_L_07_mess, data_mess$CAR_L_07_mess))

Q10_CAR_L_01_meet <- quantile(data_meet$CAR_L_01_meet, 0.01)
Q90_CAR_L_01_meet <- quantile(data_meet$CAR_L_01_meet, 0.99)
IQR_value_CAR_L_01_meet <- IQR(data_meet$CAR_L_01_meet)
lower_bound_CAR_L_01_meet <- Q10_CAR_L_01_meet - 1.5*IQR_value_CAR_L_01_meet
upper_bound_CAR_L_01_meet <- Q90_CAR_L_01_meet + 1.5*IQR_value_CAR_L_01_meet
data_meet$CAR_L_01_meet <- ifelse(data_meet$CAR_L_01_meet < lower_bound_CAR_L_01_meet, Q10_CAR_L_01_meet, ifelse(data_meet$CAR_L_01_meet > upper_bound_CAR_L_01_meet, Q90_CAR_L_01_meet, data_meet$CAR_L_01_meet))
Q10_CAR_L_01_mess <- quantile(data_mess$CAR_L_01_mess, 0.01)
Q90_CAR_L_01_mess <- quantile(data_mess$CAR_L_01_mess, 0.99)
IQR_value_CAR_L_01_mess <- IQR(data_mess$CAR_L_01_mess)
lower_bound_CAR_L_01_mess <- Q10_CAR_L_01_mess - 1.5*IQR_value_CAR_L_01_mess
upper_bound_CAR_L_01_mess <- Q90_CAR_L_01_mess + 1.5*IQR_value_CAR_L_01_mess
data_mess$CAR_L_01_mess <- ifelse(data_mess$CAR_L_01_mess < lower_bound_CAR_L_01_mess, Q10_CAR_L_01_mess, ifelse(data_mess$CAR_L_01_mess > upper_bound_CAR_L_01_mess, Q90_CAR_L_01_mess, data_mess$CAR_L_01_mess))

Q10_CAR_L_22_meet <- quantile(data_meet$CAR_L_22_meet, 0.01)
Q90_CAR_L_22_meet <- quantile(data_meet$CAR_L_22_meet, 0.99)
IQR_value_CAR_L_22_meet <- IQR(data_meet$CAR_L_22_meet)
lower_bound_CAR_L_22_meet <- Q10_CAR_L_22_meet - 1.5*IQR_value_CAR_L_22_meet
upper_bound_CAR_L_22_meet <- Q90_CAR_L_22_meet + 1.5*IQR_value_CAR_L_22_meet
data_meet$CAR_L_22_meet <- ifelse(data_meet$CAR_L_22_meet < lower_bound_CAR_L_22_meet, Q10_CAR_L_22_meet, ifelse(data_meet$CAR_L_22_meet > upper_bound_CAR_L_22_meet, Q90_CAR_L_22_meet, data_meet$CAR_L_22_meet))
Q10_CAR_L_22_mess <- quantile(data_mess$CAR_L_22_mess, 0.01)
Q90_CAR_L_22_mess <- quantile(data_mess$CAR_L_22_mess, 0.99)
IQR_value_CAR_L_22_mess <- IQR(data_mess$CAR_L_22_mess)
lower_bound_CAR_L_22_mess <- Q10_CAR_L_22_mess - 1.5*IQR_value_CAR_L_22_mess
upper_bound_CAR_L_22_mess <- Q90_CAR_L_22_mess + 1.5*IQR_value_CAR_L_22_mess
data_mess$CAR_L_22_mess <- ifelse(data_mess$CAR_L_22_mess < lower_bound_CAR_L_22_mess, Q10_CAR_L_22_mess, ifelse(data_mess$CAR_L_22_mess > upper_bound_CAR_L_22_mess, Q90_CAR_L_22_mess, data_mess$CAR_L_22_mess))

ggplot(data_meet, aes(x = CAR_L_07_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet log CAR 0 +7")
ggplot(data_mess, aes(x = CAR_L_07_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess log CAR 0 +7")
ggplot(data_meet, aes(x = CAR_L_01_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet log CAR 0 +1")
ggplot(data_mess, aes(x = CAR_L_01_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess log CAR 0 +1")
ggplot(data_meet, aes(x = CAR_L_22_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet log CAR -2 +2")
ggplot(data_mess, aes(x = CAR_L_22_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess log CAR -2 +2")


Q10_ros <- quantile(data_meet$ros, 0.1)
Q90_ros <- quantile(data_meet$ros, 0.9)
IQR_value_ros <- IQR(data_meet$ros)
lower_bound_ros <- Q10_ros - 1.5*IQR_value_ros
upper_bound_ros <- Q90_ros + 1.5*IQR_value_ros
data_meet$ros <- ifelse(data_meet$ros < lower_bound_ros, Q10_ros, ifelse(data_meet$ros > upper_bound_ros, Q90_ros, data_meet$ros))
Q10_ros <- quantile(data_mess$ros, 0.1)
Q90_ros <- quantile(data_mess$ros, 0.9)
IQR_value_ros <- IQR(data_mess$ros)
lower_bound_ros <- Q10_ros - 1.5*IQR_value_ros
upper_bound_ros <- Q90_ros + 1.5*IQR_value_ros
data_mess$ros <- ifelse(data_mess$ros < lower_bound_ros, Q10_ros, ifelse(data_mess$ros > upper_bound_ros, Q90_ros, data_mess$ros))

ggplot(data_meet, aes(x = ros)) + geom_boxplot() +
  labs(title = "Boxplot of meet ros")
ggplot(data_mess, aes(x = ros)) + geom_boxplot() +
  labs(title = "Boxplot of mess ros")

Q10_lev <- quantile(data_meet$lev, 0.1)
Q90_lev <- quantile(data_meet$lev, 0.9)
IQR_value_lev <- IQR(data_meet$lev)
lower_bound_lev <- Q10_lev - 1.5*IQR_value_lev
upper_bound_lev <- Q90_lev + 1.5*IQR_value_lev
data_meet$lev <- ifelse(data_meet$lev < lower_bound_lev, Q10_lev, ifelse(data_meet$lev > upper_bound_lev, Q90_lev, data_meet$lev))
Q10_lev <- quantile(data_mess$lev, 0.1)
Q90_lev <- quantile(data_mess$lev, 0.9)
IQR_value_lev <- IQR(data_mess$lev)
lower_bound_lev <- Q10_lev - 1.5*IQR_value_lev
upper_bound_lev <- Q90_lev + 1.5*IQR_value_lev
data_mess$lev <- ifelse(data_mess$lev < lower_bound_lev, Q10_lev, ifelse(data_mess$lev > upper_bound_lev, Q90_lev, data_mess$lev))

ggplot(data_meet, aes(x = lev)) + geom_boxplot() +
  labs(title = "Boxplot of meet lev")
ggplot(data_mess, aes(x = lev)) + geom_boxplot() +
  labs(title = "Boxplot of mess lev")

Q10_ln_act <- quantile(data_meet$ln_act, 0.1)
Q90_ln_act <- quantile(data_meet$ln_act, 0.9)
IQR_value_ln_act <- IQR(data_meet$ln_act)
lower_bound_ln_act <- Q10_ln_act - 1.5*IQR_value_ln_act
upper_bound_ln_act <- Q90_ln_act + 1.5*IQR_value_ln_act
data_meet$ln_act <- ifelse(data_meet$ln_act < lower_bound_ln_act, Q10_ln_act, ifelse(data_meet$ln_act > upper_bound_ln_act, Q90_ln_act, data_meet$ln_act))
Q10_ln_act <- quantile(data_mess$ln_act, 0.1)
Q90_ln_act <- quantile(data_mess$ln_act, 0.9)
IQR_value_ln_act <- IQR(data_mess$ln_act)
lower_bound_ln_act <- Q10_ln_act - 1.5*IQR_value_ln_act
upper_bound_ln_act <- Q90_ln_act + 1.5*IQR_value_ln_act
data_mess$ln_act <- ifelse(data_mess$ln_act < lower_bound_ln_act, Q10_ln_act, ifelse(data_mess$ln_act > upper_bound_ln_act, Q90_ln_act, data_mess$ln_act))

ggplot(data_meet, aes(x = ln_act)) + geom_boxplot() +
  labs(title = "Boxplot of meet ln_act")
ggplot(data_mess, aes(x = ln_act)) + geom_boxplot() +
  labs(title = "Boxplot of mess ln_act")

Q90_div <- quantile(data_meet$div, 0.99)
IQR_value_div <- IQR(data_meet$div)
upper_bound_div <- Q90_div + 1.5*IQR_value_div
data_meet$div <- ifelse(data_meet$div > upper_bound_div, Q90_div, data_meet$div)
Q90_div <- quantile(data_mess$div, 0.99)
IQR_value_div <- IQR(data_mess$div)
upper_bound_div <- Q90_div + 1.5*IQR_value_div
data_mess$div <- ifelse(data_mess$div > upper_bound_div, Q90_div, data_mess$div)

ggplot(data_meet, aes(x = div)) + geom_boxplot() +
  labs(title = "Boxplot of meet div")
ggplot(data_mess, aes(x = div)) + geom_boxplot() +
  labs(title = "Boxplot of mess div")


# data_meet_before2014 и data_mess_before2014

Q10_CAR_S_07_meet_before2014 <- quantile(data_meet_before2014$CAR_S_07_meet, 0.01)
Q90_CAR_S_07_meet_before2014 <- quantile(data_meet_before2014$CAR_S_07_meet, 0.99)
IQR_value_CAR_S_07_meet_before2014 <- IQR(data_meet_before2014$CAR_S_07_meet)
lower_bound_CAR_S_07_meet_before2014 <- Q10_CAR_S_07_meet_before2014 - 1.5*IQR_value_CAR_S_07_meet_before2014
upper_bound_CAR_S_07_meet_before2014 <- Q90_CAR_S_07_meet_before2014 + 1.5*IQR_value_CAR_S_07_meet_before2014
data_meet_before2014$CAR_S_07_meet <- ifelse(data_meet_before2014$CAR_S_07_meet < lower_bound_CAR_S_07_meet_before2014, Q10_CAR_S_07_meet_before2014, ifelse(data_meet_before2014$CAR_S_07_meet > upper_bound_CAR_S_07_meet_before2014, Q90_CAR_S_07_meet_before2014, data_meet_before2014$CAR_S_07_meet))
Q10_CAR_S_07_mess_before2014 <- quantile(data_mess_before2014$CAR_S_07_mess, 0.01)
Q90_CAR_S_07_mess_before2014 <- quantile(data_mess_before2014$CAR_S_07_mess, 0.99)
IQR_value_CAR_S_07_mess_before2014 <- IQR(data_mess_before2014$CAR_S_07_mess)
lower_bound_CAR_S_07_mess_before2014 <- Q10_CAR_S_07_mess_before2014 - 1.5*IQR_value_CAR_S_07_mess_before2014
upper_bound_CAR_S_07_mess_before2014 <- Q90_CAR_S_07_mess_before2014 + 1.5*IQR_value_CAR_S_07_mess_before2014
data_mess_before2014$CAR_S_07_mess <- ifelse(data_mess_before2014$CAR_S_07_mess < lower_bound_CAR_S_07_mess_before2014, Q10_CAR_S_07_mess_before2014, ifelse(data_mess_before2014$CAR_S_07_mess > upper_bound_CAR_S_07_mess_before2014, Q90_CAR_S_07_mess_before2014, data_mess_before2014$CAR_S_07_mess))

Q10_CAR_S_01_meet_before2014 <- quantile(data_meet_before2014$CAR_S_01_meet, 0.01)
Q90_CAR_S_01_meet_before2014 <- quantile(data_meet_before2014$CAR_S_01_meet, 0.99)
IQR_value_CAR_S_01_meet_before2014 <- IQR(data_meet_before2014$CAR_S_01_meet)
lower_bound_CAR_S_01_meet_before2014 <- Q10_CAR_S_01_meet_before2014 - 1.5*IQR_value_CAR_S_01_meet_before2014
upper_bound_CAR_S_01_meet_before2014 <- Q90_CAR_S_01_meet_before2014 + 1.5*IQR_value_CAR_S_01_meet_before2014
data_meet_before2014$CAR_S_01_meet <- ifelse(data_meet_before2014$CAR_S_01_meet < lower_bound_CAR_S_01_meet_before2014, Q10_CAR_S_01_meet_before2014, ifelse(data_meet_before2014$CAR_S_01_meet > upper_bound_CAR_S_01_meet_before2014, Q90_CAR_S_01_meet_before2014, data_meet_before2014$CAR_S_01_meet))
Q10_CAR_S_01_mess_before2014 <- quantile(data_mess_before2014$CAR_S_01_mess, 0.01)
Q90_CAR_S_01_mess_before2014 <- quantile(data_mess_before2014$CAR_S_01_mess, 0.99)
IQR_value_CAR_S_01_mess_before2014 <- IQR(data_mess_before2014$CAR_S_01_mess)
lower_bound_CAR_S_01_mess_before2014 <- Q10_CAR_S_01_mess_before2014 - 1.5*IQR_value_CAR_S_01_mess_before2014
upper_bound_CAR_S_01_mess_before2014 <- Q90_CAR_S_01_mess_before2014 + 1.5*IQR_value_CAR_S_01_mess_before2014
data_mess_before2014$CAR_S_01_mess <- ifelse(data_mess_before2014$CAR_S_01_mess < lower_bound_CAR_S_01_mess_before2014, Q10_CAR_S_01_mess_before2014, ifelse(data_mess_before2014$CAR_S_01_mess > upper_bound_CAR_S_01_mess_before2014, Q90_CAR_S_01_mess_before2014, data_mess_before2014$CAR_S_01_mess))

Q10_CAR_S_22_meet_before2014 <- quantile(data_meet_before2014$CAR_S_22_meet, 0.01)
Q90_CAR_S_22_meet_before2014 <- quantile(data_meet_before2014$CAR_S_22_meet, 0.99)
IQR_value_CAR_S_22_meet_before2014 <- IQR(data_meet_before2014$CAR_S_22_meet)
lower_bound_CAR_S_22_meet_before2014 <- Q10_CAR_S_22_meet_before2014 - 1.5*IQR_value_CAR_S_22_meet_before2014
upper_bound_CAR_S_22_meet_before2014 <- Q90_CAR_S_22_meet_before2014 + 1.5*IQR_value_CAR_S_22_meet_before2014
data_meet_before2014$CAR_S_22_meet <- ifelse(data_meet_before2014$CAR_S_22_meet < lower_bound_CAR_S_22_meet_before2014, Q10_CAR_S_22_meet_before2014, ifelse(data_meet_before2014$CAR_S_22_meet > upper_bound_CAR_S_22_meet_before2014, Q90_CAR_S_22_meet_before2014, data_meet_before2014$CAR_S_22_meet))
Q10_CAR_S_22_mess_before2014 <- quantile(data_mess_before2014$CAR_S_22_mess, 0.01)
Q90_CAR_S_22_mess_before2014 <- quantile(data_mess_before2014$CAR_S_22_mess, 0.99)
IQR_value_CAR_S_22_mess_before2014 <- IQR(data_mess_before2014$CAR_S_22_mess)
lower_bound_CAR_S_22_mess_before2014 <- Q10_CAR_S_22_mess_before2014 - 1.5*IQR_value_CAR_S_22_mess_before2014
upper_bound_CAR_S_22_mess_before2014 <- Q90_CAR_S_22_mess_before2014 + 1.5*IQR_value_CAR_S_22_mess_before2014
data_mess_before2014$CAR_S_22_mess <- ifelse(data_mess_before2014$CAR_S_22_mess < lower_bound_CAR_S_22_mess_before2014, Q10_CAR_S_22_mess_before2014, ifelse(data_mess_before2014$CAR_S_22_mess > upper_bound_CAR_S_22_mess_before2014, Q90_CAR_S_22_mess_before2014, data_mess_before2014$CAR_S_22_mess))

ggplot(data_meet_before2014, aes(x = CAR_S_07_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 simple CAR 0 +7")
ggplot(data_mess_before2014, aes(x = CAR_S_07_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 simple CAR 0 +7")
ggplot(data_meet_before2014, aes(x = CAR_S_01_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 simple CAR 0 +1")
ggplot(data_mess_before2014, aes(x = CAR_S_01_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 simple CAR 0 +1")
ggplot(data_meet_before2014, aes(x = CAR_S_22_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 simple CAR -2 +2")
ggplot(data_mess_before2014, aes(x = CAR_S_22_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 simple CAR -2 +2")


Q10_CAR_L_07_meet_before2014 <- quantile(data_meet_before2014$CAR_L_07_meet, 0.01)
Q90_CAR_L_07_meet_before2014 <- quantile(data_meet_before2014$CAR_L_07_meet, 0.99)
IQR_value_CAR_L_07_meet_before2014 <- IQR(data_meet_before2014$CAR_L_07_meet)
lower_bound_CAR_L_07_meet_before2014 <- Q10_CAR_L_07_meet_before2014 - 1.5*IQR_value_CAR_L_07_meet_before2014
upper_bound_CAR_L_07_meet_before2014 <- Q90_CAR_L_07_meet_before2014 + 1.5*IQR_value_CAR_L_07_meet_before2014
data_meet_before2014$CAR_L_07_meet <- ifelse(data_meet_before2014$CAR_L_07_meet < lower_bound_CAR_L_07_meet_before2014, Q10_CAR_L_07_meet_before2014, ifelse(data_meet_before2014$CAR_L_07_meet > upper_bound_CAR_L_07_meet_before2014, Q90_CAR_L_07_meet_before2014, data_meet_before2014$CAR_L_07_meet))
Q10_CAR_L_07_mess_before2014 <- quantile(data_mess_before2014$CAR_L_07_mess, 0.01)
Q90_CAR_L_07_mess_before2014 <- quantile(data_mess_before2014$CAR_L_07_mess, 0.99)
IQR_value_CAR_L_07_mess_before2014 <- IQR(data_mess_before2014$CAR_L_07_mess)
lower_bound_CAR_L_07_mess_before2014 <- Q10_CAR_L_07_mess_before2014 - 1.5*IQR_value_CAR_L_07_mess_before2014
upper_bound_CAR_L_07_mess_before2014 <- Q90_CAR_L_07_mess_before2014 + 1.5*IQR_value_CAR_L_07_mess_before2014
data_mess_before2014$CAR_L_07_mess <- ifelse(data_mess_before2014$CAR_L_07_mess < lower_bound_CAR_L_07_mess_before2014, Q10_CAR_L_07_mess_before2014, ifelse(data_mess_before2014$CAR_L_07_mess > upper_bound_CAR_L_07_mess_before2014, Q90_CAR_L_07_mess_before2014, data_mess_before2014$CAR_L_07_mess))

Q10_CAR_L_01_meet_before2014 <- quantile(data_meet_before2014$CAR_L_01_meet, 0.01)
Q90_CAR_L_01_meet_before2014 <- quantile(data_meet_before2014$CAR_L_01_meet, 0.99)
IQR_value_CAR_L_01_meet_before2014 <- IQR(data_meet_before2014$CAR_L_01_meet)
lower_bound_CAR_L_01_meet_before2014 <- Q10_CAR_L_01_meet_before2014 - 1.5*IQR_value_CAR_L_01_meet_before2014
upper_bound_CAR_L_01_meet_before2014 <- Q90_CAR_L_01_meet_before2014 + 1.5*IQR_value_CAR_L_01_meet_before2014
data_meet_before2014$CAR_L_01_meet <- ifelse(data_meet_before2014$CAR_L_01_meet < lower_bound_CAR_L_01_meet_before2014, Q10_CAR_L_01_meet_before2014, ifelse(data_meet_before2014$CAR_L_01_meet > upper_bound_CAR_L_01_meet_before2014, Q90_CAR_L_01_meet_before2014, data_meet_before2014$CAR_L_01_meet))
Q10_CAR_L_01_mess_before2014 <- quantile(data_mess_before2014$CAR_L_01_mess, 0.01)
Q90_CAR_L_01_mess_before2014 <- quantile(data_mess_before2014$CAR_L_01_mess, 0.99)
IQR_value_CAR_L_01_mess_before2014 <- IQR(data_mess_before2014$CAR_L_01_mess)
lower_bound_CAR_L_01_mess_before2014 <- Q10_CAR_L_01_mess_before2014 - 1.5*IQR_value_CAR_L_01_mess_before2014
upper_bound_CAR_L_01_mess_before2014 <- Q90_CAR_L_01_mess_before2014 + 1.5*IQR_value_CAR_L_01_mess_before2014
data_mess_before2014$CAR_L_01_mess <- ifelse(data_mess_before2014$CAR_L_01_mess < lower_bound_CAR_L_01_mess_before2014, Q10_CAR_L_01_mess_before2014, ifelse(data_mess_before2014$CAR_L_01_mess > upper_bound_CAR_L_01_mess_before2014, Q90_CAR_L_01_mess_before2014, data_mess_before2014$CAR_L_01_mess))

Q10_CAR_L_22_meet_before2014 <- quantile(data_meet_before2014$CAR_L_22_meet, 0.01)
Q90_CAR_L_22_meet_before2014 <- quantile(data_meet_before2014$CAR_L_22_meet, 0.99)
IQR_value_CAR_L_22_meet_before2014 <- IQR(data_meet_before2014$CAR_L_22_meet)
lower_bound_CAR_L_22_meet_before2014 <- Q10_CAR_L_22_meet_before2014 - 1.5*IQR_value_CAR_L_22_meet_before2014
upper_bound_CAR_L_22_meet_before2014 <- Q90_CAR_L_22_meet_before2014 + 1.5*IQR_value_CAR_L_22_meet_before2014
data_meet_before2014$CAR_L_22_meet <- ifelse(data_meet_before2014$CAR_L_22_meet < lower_bound_CAR_L_22_meet_before2014, Q10_CAR_L_22_meet_before2014, ifelse(data_meet_before2014$CAR_L_22_meet > upper_bound_CAR_L_22_meet_before2014, Q90_CAR_L_22_meet_before2014, data_meet_before2014$CAR_L_22_meet))
Q10_CAR_L_22_mess_before2014 <- quantile(data_mess_before2014$CAR_L_22_mess, 0.01)
Q90_CAR_L_22_mess_before2014 <- quantile(data_mess_before2014$CAR_L_22_mess, 0.99)
IQR_value_CAR_L_22_mess_before2014 <- IQR(data_mess_before2014$CAR_L_22_mess)
lower_bound_CAR_L_22_mess_before2014 <- Q10_CAR_L_22_mess_before2014 - 1.5*IQR_value_CAR_L_22_mess_before2014
upper_bound_CAR_L_22_mess_before2014 <- Q90_CAR_L_22_mess_before2014 + 1.5*IQR_value_CAR_L_22_mess_before2014
data_mess_before2014$CAR_L_22_mess <- ifelse(data_mess_before2014$CAR_L_22_mess < lower_bound_CAR_L_22_mess_before2014, Q10_CAR_L_22_mess_before2014, ifelse(data_mess_before2014$CAR_L_22_mess > upper_bound_CAR_L_22_mess_before2014, Q90_CAR_L_22_mess_before2014, data_mess_before2014$CAR_L_22_mess))

ggplot(data_meet_before2014, aes(x = CAR_L_07_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 log CAR 0 +7")
ggplot(data_mess_before2014, aes(x = CAR_L_07_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 log CAR 0 +7")
ggplot(data_meet_before2014, aes(x = CAR_L_01_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 log CAR 0 +1")
ggplot(data_mess_before2014, aes(x = CAR_L_01_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 log CAR 0 +1")
ggplot(data_meet_before2014, aes(x = CAR_L_22_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 log CAR -2 +2")
ggplot(data_mess_before2014, aes(x = CAR_L_22_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 log CAR -2 +2")


Q10_ros <- quantile(data_meet_before2014$ros, 0.1)
Q90_ros <- quantile(data_meet_before2014$ros, 0.9)
IQR_value_ros <- IQR(data_meet_before2014$ros)
lower_bound_ros <- Q10_ros - 1.5*IQR_value_ros
upper_bound_ros <- Q90_ros + 1.5*IQR_value_ros
data_meet_before2014$ros <- ifelse(data_meet_before2014$ros < lower_bound_ros, Q10_ros, ifelse(data_meet_before2014$ros > upper_bound_ros, Q90_ros, data_meet_before2014$ros))
Q10_ros <- quantile(data_mess_before2014$ros, 0.1)
Q90_ros <- quantile(data_mess_before2014$ros, 0.9)
IQR_value_ros <- IQR(data_mess_before2014$ros)
lower_bound_ros <- Q10_ros - 1.5*IQR_value_ros
upper_bound_ros <- Q90_ros + 1.5*IQR_value_ros
data_mess_before2014$ros <- ifelse(data_mess_before2014$ros < lower_bound_ros, Q10_ros, ifelse(data_mess_before2014$ros > upper_bound_ros, Q90_ros, data_mess_before2014$ros))

ggplot(data_meet_before2014, aes(x = ros)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 ros")
ggplot(data_mess_before2014, aes(x = ros)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 ros")

Q10_lev <- quantile(data_meet_before2014$lev, 0.1)
Q90_lev <- quantile(data_meet_before2014$lev, 0.9)
IQR_value_lev <- IQR(data_meet_before2014$lev)
lower_bound_lev <- Q10_lev - 1.5*IQR_value_lev
upper_bound_lev <- Q90_lev + 1.5*IQR_value_lev
data_meet_before2014$lev <- ifelse(data_meet_before2014$lev < lower_bound_lev, Q10_lev, ifelse(data_meet_before2014$lev > upper_bound_lev, Q90_lev, data_meet_before2014$lev))
Q10_lev <- quantile(data_mess_before2014$lev, 0.1)
Q90_lev <- quantile(data_mess_before2014$lev, 0.9)
IQR_value_lev <- IQR(data_mess_before2014$lev)
lower_bound_lev <- Q10_lev - 1.5*IQR_value_lev
upper_bound_lev <- Q90_lev + 1.5*IQR_value_lev
data_mess_before2014$lev <- ifelse(data_mess_before2014$lev < lower_bound_lev, Q10_lev, ifelse(data_mess_before2014$lev > upper_bound_lev, Q90_lev, data_mess_before2014$lev))

ggplot(data_meet_before2014, aes(x = lev)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 lev")
ggplot(data_mess_before2014, aes(x = lev)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 lev")

Q10_ln_act <- quantile(data_meet_before2014$ln_act, 0.1)
Q90_ln_act <- quantile(data_meet_before2014$ln_act, 0.9)
IQR_value_ln_act <- IQR(data_meet_before2014$ln_act)
lower_bound_ln_act <- Q10_ln_act - 1.5*IQR_value_ln_act
upper_bound_ln_act <- Q90_ln_act + 1.5*IQR_value_ln_act
data_meet_before2014$ln_act <- ifelse(data_meet_before2014$ln_act < lower_bound_ln_act, Q10_ln_act, ifelse(data_meet_before2014$ln_act > upper_bound_ln_act, Q90_ln_act, data_meet_before2014$ln_act))
Q10_ln_act <- quantile(data_mess_before2014$ln_act, 0.1)
Q90_ln_act <- quantile(data_mess_before2014$ln_act, 0.9)
IQR_value_ln_act <- IQR(data_mess_before2014$ln_act)
lower_bound_ln_act <- Q10_ln_act - 1.5*IQR_value_ln_act
upper_bound_ln_act <- Q90_ln_act + 1.5*IQR_value_ln_act
data_mess_before2014$ln_act <- ifelse(data_mess_before2014$ln_act < lower_bound_ln_act, Q10_ln_act, ifelse(data_mess_before2014$ln_act > upper_bound_ln_act, Q90_ln_act, data_mess_before2014$ln_act))

ggplot(data_meet_before2014, aes(x = ln_act)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 ln_act")
ggplot(data_mess_before2014, aes(x = ln_act)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 ln_act")

Q90_div <- quantile(data_meet_before2014$div, 0.99)
IQR_value_div <- IQR(data_meet_before2014$div)
upper_bound_div <- Q90_div + 1.5*IQR_value_div
data_meet_before2014$div <- ifelse(data_meet_before2014$div > upper_bound_div, Q90_div, data_meet_before2014$div)
Q90_div <- quantile(data_mess_before2014$div, 0.99)
IQR_value_div <- IQR(data_mess_before2014$div)
upper_bound_div <- Q90_div + 1.5*IQR_value_div
data_mess_before2014$div <- ifelse(data_mess_before2014$div > upper_bound_div, Q90_div, data_mess_before2014$div)

ggplot(data_meet_before2014, aes(x = div)) + geom_boxplot() +
  labs(title = "Boxplot of meet_before2014 div")
ggplot(data_mess_before2014, aes(x = div)) + geom_boxplot() +
  labs(title = "Boxplot of mess_before2014 div")

# data_meet_after2014 и data_mess_after2014

Q10_CAR_S_07_meet_after2014 <- quantile(data_meet_after2014$CAR_S_07_meet, 0.01)
Q90_CAR_S_07_meet_after2014 <- quantile(data_meet_after2014$CAR_S_07_meet, 0.99)
IQR_value_CAR_S_07_meet_after2014 <- IQR(data_meet_after2014$CAR_S_07_meet)
lower_bound_CAR_S_07_meet_after2014 <- Q10_CAR_S_07_meet_after2014 - 1.5*IQR_value_CAR_S_07_meet_after2014
upper_bound_CAR_S_07_meet_after2014 <- Q90_CAR_S_07_meet_after2014 + 1.5*IQR_value_CAR_S_07_meet_after2014
data_meet_after2014$CAR_S_07_meet <- ifelse(data_meet_after2014$CAR_S_07_meet < lower_bound_CAR_S_07_meet_after2014, Q10_CAR_S_07_meet_after2014, ifelse(data_meet_after2014$CAR_S_07_meet > upper_bound_CAR_S_07_meet_after2014, Q90_CAR_S_07_meet_after2014, data_meet_after2014$CAR_S_07_meet))
Q10_CAR_S_07_mess_after2014 <- quantile(data_mess_after2014$CAR_S_07_mess, 0.01)
Q90_CAR_S_07_mess_after2014 <- quantile(data_mess_after2014$CAR_S_07_mess, 0.99)
IQR_value_CAR_S_07_mess_after2014 <- IQR(data_mess_after2014$CAR_S_07_mess)
lower_bound_CAR_S_07_mess_after2014 <- Q10_CAR_S_07_mess_after2014 - 1.5*IQR_value_CAR_S_07_mess_after2014
upper_bound_CAR_S_07_mess_after2014 <- Q90_CAR_S_07_mess_after2014 + 1.5*IQR_value_CAR_S_07_mess_after2014
data_mess_after2014$CAR_S_07_mess <- ifelse(data_mess_after2014$CAR_S_07_mess < lower_bound_CAR_S_07_mess_after2014, Q10_CAR_S_07_mess_after2014, ifelse(data_mess_after2014$CAR_S_07_mess > upper_bound_CAR_S_07_mess_after2014, Q90_CAR_S_07_mess_after2014, data_mess_after2014$CAR_S_07_mess))

Q10_CAR_S_01_meet_after2014 <- quantile(data_meet_after2014$CAR_S_01_meet, 0.01)
Q90_CAR_S_01_meet_after2014 <- quantile(data_meet_after2014$CAR_S_01_meet, 0.99)
IQR_value_CAR_S_01_meet_after2014 <- IQR(data_meet_after2014$CAR_S_01_meet)
lower_bound_CAR_S_01_meet_after2014 <- Q10_CAR_S_01_meet_after2014 - 1.5*IQR_value_CAR_S_01_meet_after2014
upper_bound_CAR_S_01_meet_after2014 <- Q90_CAR_S_01_meet_after2014 + 1.5*IQR_value_CAR_S_01_meet_after2014
data_meet_after2014$CAR_S_01_meet <- ifelse(data_meet_after2014$CAR_S_01_meet < lower_bound_CAR_S_01_meet_after2014, Q10_CAR_S_01_meet_after2014, ifelse(data_meet_after2014$CAR_S_01_meet > upper_bound_CAR_S_01_meet_after2014, Q90_CAR_S_01_meet_after2014, data_meet_after2014$CAR_S_01_meet))
Q10_CAR_S_01_mess_after2014 <- quantile(data_mess_after2014$CAR_S_01_mess, 0.01)
Q90_CAR_S_01_mess_after2014 <- quantile(data_mess_after2014$CAR_S_01_mess, 0.99)
IQR_value_CAR_S_01_mess_after2014 <- IQR(data_mess_after2014$CAR_S_01_mess)
lower_bound_CAR_S_01_mess_after2014 <- Q10_CAR_S_01_mess_after2014 - 1.5*IQR_value_CAR_S_01_mess_after2014
upper_bound_CAR_S_01_mess_after2014 <- Q90_CAR_S_01_mess_after2014 + 1.5*IQR_value_CAR_S_01_mess_after2014
data_mess_after2014$CAR_S_01_mess <- ifelse(data_mess_after2014$CAR_S_01_mess < lower_bound_CAR_S_01_mess_after2014, Q10_CAR_S_01_mess_after2014, ifelse(data_mess_after2014$CAR_S_01_mess > upper_bound_CAR_S_01_mess_after2014, Q90_CAR_S_01_mess_after2014, data_mess_after2014$CAR_S_01_mess))

Q10_CAR_S_22_meet_after2014 <- quantile(data_meet_after2014$CAR_S_22_meet, 0.01)
Q90_CAR_S_22_meet_after2014 <- quantile(data_meet_after2014$CAR_S_22_meet, 0.99)
IQR_value_CAR_S_22_meet_after2014 <- IQR(data_meet_after2014$CAR_S_22_meet)
lower_bound_CAR_S_22_meet_after2014 <- Q10_CAR_S_22_meet_after2014 - 1.5*IQR_value_CAR_S_22_meet_after2014
upper_bound_CAR_S_22_meet_after2014 <- Q90_CAR_S_22_meet_after2014 + 1.5*IQR_value_CAR_S_22_meet_after2014
data_meet_after2014$CAR_S_22_meet <- ifelse(data_meet_after2014$CAR_S_22_meet < lower_bound_CAR_S_22_meet_after2014, Q10_CAR_S_22_meet_after2014, ifelse(data_meet_after2014$CAR_S_22_meet > upper_bound_CAR_S_22_meet_after2014, Q90_CAR_S_22_meet_after2014, data_meet_after2014$CAR_S_22_meet))
Q10_CAR_S_22_mess_after2014 <- quantile(data_mess_after2014$CAR_S_22_mess, 0.01)
Q90_CAR_S_22_mess_after2014 <- quantile(data_mess_after2014$CAR_S_22_mess, 0.99)
IQR_value_CAR_S_22_mess_after2014 <- IQR(data_mess_after2014$CAR_S_22_mess)
lower_bound_CAR_S_22_mess_after2014 <- Q10_CAR_S_22_mess_after2014 - 1.5*IQR_value_CAR_S_22_mess_after2014
upper_bound_CAR_S_22_mess_after2014 <- Q90_CAR_S_22_mess_after2014 + 1.5*IQR_value_CAR_S_22_mess_after2014
data_mess_after2014$CAR_S_22_mess <- ifelse(data_mess_after2014$CAR_S_22_mess < lower_bound_CAR_S_22_mess_after2014, Q10_CAR_S_22_mess_after2014, ifelse(data_mess_after2014$CAR_S_22_mess > upper_bound_CAR_S_22_mess_after2014, Q90_CAR_S_22_mess_after2014, data_mess_after2014$CAR_S_22_mess))

ggplot(data_meet_after2014, aes(x = CAR_S_07_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 simple CAR 0 +7")
ggplot(data_mess_after2014, aes(x = CAR_S_07_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 simple CAR 0 +7")
ggplot(data_meet_after2014, aes(x = CAR_S_01_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 simple CAR 0 +1")
ggplot(data_mess_after2014, aes(x = CAR_S_01_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 simple CAR 0 +1")
ggplot(data_meet_after2014, aes(x = CAR_S_22_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 simple CAR -2 +2")
ggplot(data_mess_after2014, aes(x = CAR_S_22_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 simple CAR -2 +2")


Q10_CAR_L_07_meet_after2014 <- quantile(data_meet_after2014$CAR_L_07_meet, 0.01)
Q90_CAR_L_07_meet_after2014 <- quantile(data_meet_after2014$CAR_L_07_meet, 0.99)
IQR_value_CAR_L_07_meet_after2014 <- IQR(data_meet_after2014$CAR_L_07_meet)
lower_bound_CAR_L_07_meet_after2014 <- Q10_CAR_L_07_meet_after2014 - 1.5*IQR_value_CAR_L_07_meet_after2014
upper_bound_CAR_L_07_meet_after2014 <- Q90_CAR_L_07_meet_after2014 + 1.5*IQR_value_CAR_L_07_meet_after2014
data_meet_after2014$CAR_L_07_meet <- ifelse(data_meet_after2014$CAR_L_07_meet < lower_bound_CAR_L_07_meet_after2014, Q10_CAR_L_07_meet_after2014, ifelse(data_meet_after2014$CAR_L_07_meet > upper_bound_CAR_L_07_meet_after2014, Q90_CAR_L_07_meet_after2014, data_meet_after2014$CAR_L_07_meet))
Q10_CAR_L_07_mess_after2014 <- quantile(data_mess_after2014$CAR_L_07_mess, 0.01)
Q90_CAR_L_07_mess_after2014 <- quantile(data_mess_after2014$CAR_L_07_mess, 0.99)
IQR_value_CAR_L_07_mess_after2014 <- IQR(data_mess_after2014$CAR_L_07_mess)
lower_bound_CAR_L_07_mess_after2014 <- Q10_CAR_L_07_mess_after2014 - 1.5*IQR_value_CAR_L_07_mess_after2014
upper_bound_CAR_L_07_mess_after2014 <- Q90_CAR_L_07_mess_after2014 + 1.5*IQR_value_CAR_L_07_mess_after2014
data_mess_after2014$CAR_L_07_mess <- ifelse(data_mess_after2014$CAR_L_07_mess < lower_bound_CAR_L_07_mess_after2014, Q10_CAR_L_07_mess_after2014, ifelse(data_mess_after2014$CAR_L_07_mess > upper_bound_CAR_L_07_mess_after2014, Q90_CAR_L_07_mess_after2014, data_mess_after2014$CAR_L_07_mess))

Q10_CAR_L_01_meet_after2014 <- quantile(data_meet_after2014$CAR_L_01_meet, 0.01)
Q90_CAR_L_01_meet_after2014 <- quantile(data_meet_after2014$CAR_L_01_meet, 0.99)
IQR_value_CAR_L_01_meet_after2014 <- IQR(data_meet_after2014$CAR_L_01_meet)
lower_bound_CAR_L_01_meet_after2014 <- Q10_CAR_L_01_meet_after2014 - 1.5*IQR_value_CAR_L_01_meet_after2014
upper_bound_CAR_L_01_meet_after2014 <- Q90_CAR_L_01_meet_after2014 + 1.5*IQR_value_CAR_L_01_meet_after2014
data_meet_after2014$CAR_L_01_meet <- ifelse(data_meet_after2014$CAR_L_01_meet < lower_bound_CAR_L_01_meet_after2014, Q10_CAR_L_01_meet_after2014, ifelse(data_meet_after2014$CAR_L_01_meet > upper_bound_CAR_L_01_meet_after2014, Q90_CAR_L_01_meet_after2014, data_meet_after2014$CAR_L_01_meet))
Q10_CAR_L_01_mess_after2014 <- quantile(data_mess_after2014$CAR_L_01_mess, 0.01)
Q90_CAR_L_01_mess_after2014 <- quantile(data_mess_after2014$CAR_L_01_mess, 0.99)
IQR_value_CAR_L_01_mess_after2014 <- IQR(data_mess_after2014$CAR_L_01_mess)
lower_bound_CAR_L_01_mess_after2014 <- Q10_CAR_L_01_mess_after2014 - 1.5*IQR_value_CAR_L_01_mess_after2014
upper_bound_CAR_L_01_mess_after2014 <- Q90_CAR_L_01_mess_after2014 + 1.5*IQR_value_CAR_L_01_mess_after2014
data_mess_after2014$CAR_L_01_mess <- ifelse(data_mess_after2014$CAR_L_01_mess < lower_bound_CAR_L_01_mess_after2014, Q10_CAR_L_01_mess_after2014, ifelse(data_mess_after2014$CAR_L_01_mess > upper_bound_CAR_L_01_mess_after2014, Q90_CAR_L_01_mess_after2014, data_mess_after2014$CAR_L_01_mess))

Q10_CAR_L_22_meet_after2014 <- quantile(data_meet_after2014$CAR_L_22_meet, 0.01)
Q90_CAR_L_22_meet_after2014 <- quantile(data_meet_after2014$CAR_L_22_meet, 0.99)
IQR_value_CAR_L_22_meet_after2014 <- IQR(data_meet_after2014$CAR_L_22_meet)
lower_bound_CAR_L_22_meet_after2014 <- Q10_CAR_L_22_meet_after2014 - 1.5*IQR_value_CAR_L_22_meet_after2014
upper_bound_CAR_L_22_meet_after2014 <- Q90_CAR_L_22_meet_after2014 + 1.5*IQR_value_CAR_L_22_meet_after2014
data_meet_after2014$CAR_L_22_meet <- ifelse(data_meet_after2014$CAR_L_22_meet < lower_bound_CAR_L_22_meet_after2014, Q10_CAR_L_22_meet_after2014, ifelse(data_meet_after2014$CAR_L_22_meet > upper_bound_CAR_L_22_meet_after2014, Q90_CAR_L_22_meet_after2014, data_meet_after2014$CAR_L_22_meet))
Q10_CAR_L_22_mess_after2014 <- quantile(data_mess_after2014$CAR_L_22_mess, 0.01)
Q90_CAR_L_22_mess_after2014 <- quantile(data_mess_after2014$CAR_L_22_mess, 0.99)
IQR_value_CAR_L_22_mess_after2014 <- IQR(data_mess_after2014$CAR_L_22_mess)
lower_bound_CAR_L_22_mess_after2014 <- Q10_CAR_L_22_mess_after2014 - 1.5*IQR_value_CAR_L_22_mess_after2014
upper_bound_CAR_L_22_mess_after2014 <- Q90_CAR_L_22_mess_after2014 + 1.5*IQR_value_CAR_L_22_mess_after2014
data_mess_after2014$CAR_L_22_mess <- ifelse(data_mess_after2014$CAR_L_22_mess < lower_bound_CAR_L_22_mess_after2014, Q10_CAR_L_22_mess_after2014, ifelse(data_mess_after2014$CAR_L_22_mess > upper_bound_CAR_L_22_mess_after2014, Q90_CAR_L_22_mess_after2014, data_mess_after2014$CAR_L_22_mess))

ggplot(data_meet_after2014, aes(x = CAR_L_07_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 log CAR 0 +7")
ggplot(data_mess_after2014, aes(x = CAR_L_07_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 log CAR 0 +7")
ggplot(data_meet_after2014, aes(x = CAR_L_01_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 log CAR 0 +1")
ggplot(data_mess_after2014, aes(x = CAR_L_01_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 log CAR 0 +1")
ggplot(data_meet_after2014, aes(x = CAR_L_22_meet)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 log CAR -2 +2")
ggplot(data_mess_after2014, aes(x = CAR_L_22_mess)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 log CAR -2 +2")


Q10_ros <- quantile(data_meet_after2014$ros, 0.1)
Q90_ros <- quantile(data_meet_after2014$ros, 0.9)
IQR_value_ros <- IQR(data_meet_after2014$ros)
lower_bound_ros <- Q10_ros - 1.5*IQR_value_ros
upper_bound_ros <- Q90_ros + 1.5*IQR_value_ros
data_meet_after2014$ros <- ifelse(data_meet_after2014$ros < lower_bound_ros, Q10_ros, ifelse(data_meet_after2014$ros > upper_bound_ros, Q90_ros, data_meet_after2014$ros))
Q10_ros <- quantile(data_mess_after2014$ros, 0.1)
Q90_ros <- quantile(data_mess_after2014$ros, 0.9)
IQR_value_ros <- IQR(data_mess_after2014$ros)
lower_bound_ros <- Q10_ros - 1.5*IQR_value_ros
upper_bound_ros <- Q90_ros + 1.5*IQR_value_ros
data_mess_after2014$ros <- ifelse(data_mess_after2014$ros < lower_bound_ros, Q10_ros, ifelse(data_mess_after2014$ros > upper_bound_ros, Q90_ros, data_mess_after2014$ros))

ggplot(data_meet_after2014, aes(x = ros)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 ros")
ggplot(data_mess_after2014, aes(x = ros)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 ros")

Q10_lev <- quantile(data_meet_after2014$lev, 0.1)
Q90_lev <- quantile(data_meet_after2014$lev, 0.9)
IQR_value_lev <- IQR(data_meet_after2014$lev)
lower_bound_lev <- Q10_lev - 1.5*IQR_value_lev
upper_bound_lev <- Q90_lev + 1.5*IQR_value_lev
data_meet_after2014$lev <- ifelse(data_meet_after2014$lev < lower_bound_lev, Q10_lev, ifelse(data_meet_after2014$lev > upper_bound_lev, Q90_lev, data_meet_after2014$lev))
Q10_lev <- quantile(data_mess_after2014$lev, 0.1)
Q90_lev <- quantile(data_mess_after2014$lev, 0.9)
IQR_value_lev <- IQR(data_mess_after2014$lev)
lower_bound_lev <- Q10_lev - 1.5*IQR_value_lev
upper_bound_lev <- Q90_lev + 1.5*IQR_value_lev
data_mess_after2014$lev <- ifelse(data_mess_after2014$lev < lower_bound_lev, Q10_lev, ifelse(data_mess_after2014$lev > upper_bound_lev, Q90_lev, data_mess_after2014$lev))

ggplot(data_meet_after2014, aes(x = lev)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 lev")
ggplot(data_mess_after2014, aes(x = lev)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 lev")

Q10_ln_act <- quantile(data_meet_after2014$ln_act, 0.1)
Q90_ln_act <- quantile(data_meet_after2014$ln_act, 0.9)
IQR_value_ln_act <- IQR(data_meet_after2014$ln_act)
lower_bound_ln_act <- Q10_ln_act - 1.5*IQR_value_ln_act
upper_bound_ln_act <- Q90_ln_act + 1.5*IQR_value_ln_act
data_meet_after2014$ln_act <- ifelse(data_meet_after2014$ln_act < lower_bound_ln_act, Q10_ln_act, ifelse(data_meet_after2014$ln_act > upper_bound_ln_act, Q90_ln_act, data_meet_after2014$ln_act))
Q10_ln_act <- quantile(data_mess_after2014$ln_act, 0.1)
Q90_ln_act <- quantile(data_mess_after2014$ln_act, 0.9)
IQR_value_ln_act <- IQR(data_mess_after2014$ln_act)
lower_bound_ln_act <- Q10_ln_act - 1.5*IQR_value_ln_act
upper_bound_ln_act <- Q90_ln_act + 1.5*IQR_value_ln_act
data_mess_after2014$ln_act <- ifelse(data_mess_after2014$ln_act < lower_bound_ln_act, Q10_ln_act, ifelse(data_mess_after2014$ln_act > upper_bound_ln_act, Q90_ln_act, data_mess_after2014$ln_act))

ggplot(data_meet_after2014, aes(x = ln_act)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 ln_act")
ggplot(data_mess_after2014, aes(x = ln_act)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 ln_act")

Q90_div <- quantile(data_meet_after2014$div, 0.99)
IQR_value_div <- IQR(data_meet_after2014$div)
upper_bound_div <- Q90_div + 1.5*IQR_value_div
data_meet_after2014$div <- ifelse(data_meet_after2014$div > upper_bound_div, Q90_div, data_meet_after2014$div)
Q90_div <- quantile(data_mess_after2014$div, 0.99)
IQR_value_div <- IQR(data_mess_after2014$div)
upper_bound_div <- Q90_div + 1.5*IQR_value_div
data_mess_after2014$div <- ifelse(data_mess_after2014$div > upper_bound_div, Q90_div, data_mess_after2014$div)

ggplot(data_meet_after2014, aes(x = div)) + geom_boxplot() +
  labs(title = "Boxplot of meet_after2014 div")
ggplot(data_mess_after2014, aes(x = div)) + geom_boxplot() +
  labs(title = "Boxplot of mess_after2014 div")

# !!!Гистограммы!!!

# data_meet и data_mess

ggplot(data_meet, aes(x=CAR_L_07_meet)) + geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black") + labs(x="CAR_L_07_meet", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=CAR_L_07_mess)) + geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black") + labs(x="CAR_L_07_mess", y = "Частота") + theme_minimal()
ggplot(data_meet, aes(x=CAR_L_01_meet)) + geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black") + labs(x="CAR_L_01_meet", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=CAR_L_01_mess)) + geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black") + labs(x="CAR_L_01_mess", y = "Частота") + theme_minimal()
ggplot(data_meet, aes(x=CAR_L_22_meet)) + geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black") + labs(x="CAR_L_22_meet", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=CAR_L_22_mess)) + geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black") + labs(x="CAR_L_22_mess", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=ln_act)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="ln_act", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=ln_act)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="ln_act", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=lev)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="lev", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=lev)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="lev", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=ros)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="ros", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=ros)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="ros", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=gov_prop)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="gov_prop", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=gov_prop)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="gov_prop", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=div)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="div решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=div)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="div голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=share_cap_diff)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="share_cap_diff", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=share_cap_diff)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="share_cap_diff", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=mer_acq)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="mer_acq решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=mer_acq)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="mer_acq голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=factor(pre_board))) + geom_bar(fill = "steelblue", color = "black") + labs(x="pre_board решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=factor(pre_board))) + geom_bar(fill = "steelblue", color = "black") + labs(x="pre_board голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=factor(diff_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="diff_gen решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=factor(diff_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="diff_gen голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=strong_current)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="strong_current решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=strong_current)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="strong_current голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=strong_exp)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="strong_exp решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=strong_exp)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="strong_exp голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=weak_current)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="weak_current решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=weak_current)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="weak_current голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=weak_exp)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="weak_exp решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=weak_exp)) + geom_histogram(fill = "steelblue", color = "black") + labs(x="weak_exp голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=factor(strong_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="strong_gen решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=factor(strong_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="strong_gen голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=factor(weak_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="weak_gen решения", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=factor(weak_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="weak_gen голосование", y = "Частота") + theme_minimal()

ggplot(data_meet, aes(x=factor(polit_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="polit_gen", y = "Частота") + theme_minimal()
ggplot(data_mess, aes(x=factor(polit_gen))) + geom_bar(fill = "steelblue", color = "black") + labs(x="polit_gen", y = "Частота") + theme_minimal()

# матрицы корреляций

# data_meet

# 07

cor_matrix_L_07_meet_current <- cor(data_meet[, c("CAR_L_07_meet", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_07_meet_current)
corrplot(cor_matrix_L_07_meet_current)
cor_matrix_L_07_meet_exp <- cor(data_meet[, c("CAR_L_07_meet", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_07_meet_exp)
corrplot(cor_matrix_L_07_meet_exp)

# 01

cor_matrix_L_01_meet_current <- cor(data_meet[, c("CAR_L_01_meet", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_01_meet_current)
corrplot(cor_matrix_L_01_meet_current)
cor_matrix_L_01_meet_exp <- cor(data_meet[, c("CAR_L_01_meet", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_01_meet_exp)
corrplot(cor_matrix_L_01_meet_exp)

# -2+2

cor_matrix_L_22_meet_current <- cor(data_meet[, c("CAR_L_22_meet", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_22_meet_current)
corrplot(cor_matrix_L_22_meet_current)
cor_matrix_L_22_meet_exp <- cor(data_meet[, c("CAR_L_22_meet", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_22_meet_exp)
corrplot(cor_matrix_L_22_meet_exp)


# data_mess

# 07

cor_matrix_L_07_mess_current <- cor(data_mess[, c("CAR_L_07_mess", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_07_mess_current)
corrplot(cor_matrix_L_07_mess_current)
cor_matrix_L_07_mess_exp <- cor(data_mess[, c("CAR_L_07_mess", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_07_mess_exp)
corrplot(cor_matrix_L_07_mess_exp)

# 01

cor_matrix_L_01_mess_current <- cor(data_mess[, c("CAR_L_01_mess", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_01_mess_current)
corrplot(cor_matrix_L_01_mess_current)
cor_matrix_L_01_mess_exp <- cor(data_mess[, c("CAR_L_01_mess", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_01_mess_exp)
corrplot(cor_matrix_L_01_mess_exp)

# -2+2

cor_matrix_L_22_mess_current <- cor(data_mess[, c("CAR_L_22_mess", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_22_mess_current)
corrplot(cor_matrix_L_22_mess_current)
cor_matrix_L_22_mess_exp <- cor(data_mess[, c("CAR_L_22_mess", "ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")])
print(cor_matrix_L_22_mess_exp)
corrplot(cor_matrix_L_22_mess_exp)

#!!!
# МОДЕЛИ
#!!!

# Все собрания

# Решения ОКНО от 0 до +7

meet_L_07_all_current <- lm(data = data_meet, CAR_L_07_meet ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_L_07_all_current)
meet_L_07_all_exp <- lm(data = data_meet, CAR_L_07_meet ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_L_07_all_exp)

#!!!#
stargazer(meet_L_07_all_current, type = "text")
stargazer(meet_L_07_all_exp, type = "text")
#!!!#

# Рассылка ОКНО от 0 до +7

mess_L_07_gov0_cur <- lm(
  CAR_L_07_mess ~ ln_act + lev + ros + div + share_cap_diff + 
    mer_acq + pre_board + diff_gen + strong_current + weak_current + 
    strong_gen + weak_gen + polit_gen,
  data = subset(data_mess, gov_prop == 0)
)

# Модель для наблюдений, где gov_prop != 0
mess_L_07_gov1_cur <- lm(
  CAR_L_07_mess ~ ln_act + lev + ros + div + share_cap_diff + 
    mer_acq + pre_board + diff_gen + strong_current + weak_current + 
    strong_gen + weak_gen + polit_gen,
  data = subset(data_mess, gov_prop != 0)
)

# Вывод обеих таблиц
summary(mess_L_07_gov0_cur)
summary(mess_L_07_gov1_cur)

mess_L_07_all_exp <- lm(data = data_mess, CAR_L_07_mess ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_L_07_all_exp)

#!!!#
stargazer(mess_L_07_all_current, type = "text")
stargazer(mess_L_07_all_exp, type = "text")
#!!!#

# Решения ОКНО от 0 до +1

meet_L_01_all_current <- lm(data = data_meet, CAR_L_01_meet ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_L_01_all_current)
meet_L_01_all_exp <- lm(data = data_meet, CAR_L_01_meet ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_L_01_all_exp)

#!!!#
stargazer(meet_L_01_all_current, type = "text")
stargazer(meet_L_01_all_exp, type = "text")
#!!!#

# Рассылка ОКНО от 0 до +1

mess_L_01_all_current <- lm(data = data_mess, CAR_L_01_mess ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_L_01_all_current)
mess_L_01_all_exp <- lm(data = data_mess, CAR_L_01_mess ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_L_01_all_exp)

#!!!#
stargazer(mess_L_01_all_current, type = "text")
stargazer(mess_L_01_all_exp, type = "text")
#!!!#

# Решения ОКНО от -2 до +2

meet_L_22_all_current <- lm(data = data_meet, CAR_L_22_meet ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_L_22_all_current)
meet_L_22_all_exp <- lm(data = data_meet, CAR_L_22_meet ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_L_22_all_exp)

#!!!#
stargazer(meet_L_22_all_current, type = "text")
stargazer(meet_L_22_all_exp, type = "text")
#!!!#

# Рассылка ОКНО от -2 до +2

mess_L_22_all_current <- lm(data = data_mess, CAR_L_22_mess ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_L_22_all_current)
mess_L_22_all_exp <- lm(data = data_mess, CAR_L_22_mess ~ ln_act + lev + ros +gov_prop + div + share_cap_diff + mer_acq + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_L_22_all_exp)

#!!!#
stargazer(mess_L_22_all_current, type = "text")
stargazer(mess_L_22_all_exp, type = "text")
#!!!#

# 1. Проверка спецификации (квадраты и кубы) resettest нулевая гипотеза все правильно, т.е если p-value>0.05, то не отклоняем, все ок
#   Текущая связь собрания
resettest(meet_L_01_all_current, power = 2:3, type = "regressor") 
resettest(meet_L_07_all_current, power = 2:3, type = "regressor")
resettest(meet_L_22_all_current, power = 2:3, type = "regressor")
#   Текущая связь рассылка
resettest(mess_L_01_all_current, power = 2:3, type = "regressor") 
resettest(mess_L_07_all_current, power = 2:3, type = "regressor")
resettest(mess_L_22_all_current, power = 2:3, type = "regressor")
#   Опыт собрания
resettest(meet_L_01_all_exp, power = 2:3, type = "regressor") 
resettest(meet_L_07_all_exp, power = 2:3, type = "regressor")
resettest(meet_L_22_all_exp, power = 2:3, type = "regressor")
#   Опыт рассылка
resettest(mess_L_01_all_exp, power = 2:3, type = "regressor") 
resettest(mess_L_07_all_exp, power = 2:3, type = "regressor")
resettest(mess_L_22_all_exp, power = 2:3, type = "regressor")
# Вывод, квадраты и кубы не нужны нигде


# 2. Равенство математических ожиданий ошибок нулю (иначе смещенный результат)
#   Нулевая гипотеза что ошибки 0, если p-value > 0.05 не отклоняем, все ок
#   Текущая связь собрания
residuals_meet_L_01_all_current <- resid(meet_L_01_all_current) 
t_test_meet_L_01_all_current <- t.test(residuals_meet_L_01_all_current, mu = 0)
print(t_test_meet_L_01_all_current)
residuals_meet_L_07_all_current <- resid(meet_L_07_all_current) 
t_test_meet_L_07_all_current <- t.test(residuals_meet_L_07_all_current, mu = 0)
print(t_test_meet_L_07_all_current)
residuals_meet_L_22_all_current <- resid(meet_L_22_all_current) 
t_test_meet_L_22_all_current <- t.test(residuals_meet_L_22_all_current, mu = 0)
print(t_test_meet_L_22_all_current)
#   Текущая связь рассылка
residuals_mess_L_01_all_current <- resid(mess_L_01_all_current) 
t_test_mess_L_01_all_current <- t.test(residuals_mess_L_01_all_current, mu = 0)
print(t_test_mess_L_01_all_current)
residuals_mess_L_07_all_current <- resid(mess_L_07_all_current) 
t_test_mess_L_07_all_current <- t.test(residuals_mess_L_07_all_current, mu = 0)
print(t_test_mess_L_07_all_current)
residuals_mess_L_22_all_current <- resid(mess_L_22_all_current) 
t_test_mess_L_22_all_current <- t.test(residuals_mess_L_22_all_current, mu = 0)
print(t_test_mess_L_22_all_current)
#   Опыт собрания
residuals_meet_L_01_all_exp <- resid(meet_L_01_all_exp) 
t_test_meet_L_01_all_exp <- t.test(residuals_meet_L_01_all_exp, mu = 0)
print(t_test_meet_L_01_all_exp)
residuals_meet_L_07_all_exp <- resid(meet_L_07_all_exp) 
t_test_meet_L_07_all_exp <- t.test(residuals_meet_L_07_all_exp, mu = 0)
print(t_test_meet_L_07_all_exp)
residuals_meet_L_22_all_exp <- resid(meet_L_22_all_exp) 
t_test_meet_L_22_all_exp <- t.test(residuals_meet_L_22_all_exp, mu = 0)
print(t_test_meet_L_22_all_exp)
#   Опыт рассылка
residuals_mess_L_01_all_exp <- resid(mess_L_01_all_exp) 
t_test_mess_L_01_all_exp <- t.test(residuals_mess_L_01_all_exp, mu = 0)
print(t_test_mess_L_01_all_exp)
residuals_mess_L_07_all_exp <- resid(mess_L_07_all_exp) 
t_test_mess_L_07_all_exp <- t.test(residuals_mess_L_07_all_exp, mu = 0)
print(t_test_mess_L_07_all_exp)
residuals_mess_L_22_all_exp <- resid(mess_L_22_all_exp) 
t_test_mess_L_22_all_exp <- t.test(residuals_mess_L_22_all_exp, mu = 0)
print(t_test_mess_L_22_all_exp)
# Вывод все ок, ошибки не отличаются от нуля, результат несмещенный

# 3. Проверка гомоскедастичности (одинаковая дисперсия остатков)
# bptest нулевая гипотеза все хорошо гомоскедастичность, если p-value > 0.05 не отклоняем, все ок
#   Текущая связь собрания
bptest(meet_L_01_all_current)
bptest(meet_L_07_all_current)
bptest(meet_L_22_all_current)
#   Текущая связь рассылка
bptest(mess_L_01_all_current)
bptest(mess_L_07_all_current)
bptest(mess_L_22_all_current)
#   Опыт собрания
bptest(meet_L_01_all_exp)
bptest(meet_L_07_all_exp)
bptest(meet_L_22_all_exp)
#   Опыт рассылка
bptest(mess_L_01_all_exp)
bptest(mess_L_07_all_exp)
bptest(mess_L_22_all_exp)
# Вывод везде гетероскедастичность, надо что-то делать или можно ничего не делать

# 4.Проверка автокорреляции ошибок
# тест дарбина уотсона, нулевая гипотеза автокорреляции ПЕРВОГО ПОРЯДКА нет, если p-value > 0.05 все ок

#   Текущая связь собрания
dw_test_meet_L_01_all_current <- dwtest(meet_L_01_all_current)
print(dw_test_meet_L_01_all_current)
dw_test_meet_L_07_all_current <- dwtest(meet_L_07_all_current)
print(dw_test_meet_L_07_all_current)
dw_test_meet_L_22_all_current <- dwtest(meet_L_22_all_current)
print(dw_test_meet_L_22_all_current)
#   Текущая связь рассылка
dw_test_mess_L_01_all_current <- dwtest(mess_L_01_all_current)
print(dw_test_mess_L_01_all_current)
dw_test_mess_L_07_all_current <- dwtest(mess_L_07_all_current)
print(dw_test_mess_L_07_all_current)
dw_test_mess_L_22_all_current <- dwtest(mess_L_22_all_current)
print(dw_test_mess_L_22_all_current)
#   Опыт собрания
dw_test_meet_L_01_all_exp <- dwtest(meet_L_01_all_exp)
print(dw_test_meet_L_01_all_exp)
dw_test_meet_L_07_all_exp <- dwtest(meet_L_07_all_exp)
print(dw_test_meet_L_07_all_exp)
dw_test_meet_L_22_all_exp <- dwtest(meet_L_22_all_exp)
print(dw_test_meet_L_22_all_exp)
#   Опыт рассылка
dw_test_mess_L_01_all_exp <- dwtest(mess_L_01_all_exp)
print(dw_test_mess_L_01_all_exp)
dw_test_mess_L_07_all_exp <- dwtest(mess_L_07_all_exp)
print(dw_test_mess_L_07_all_exp)
dw_test_mess_L_22_all_exp <- dwtest(mess_L_22_all_exp)
print(dw_test_mess_L_22_all_exp)
# Вывод L_01 не очень, остальное ок

# 5. Проверка мультиколлинеарности
# vif

#   Текущая связь собрания
vif(meet_L_01_all_current)
vif(meet_L_07_all_current)
vif(meet_L_22_all_current)
#   Текущая связь рассылка
vif(mess_L_01_all_current)
vif(mess_L_07_all_current)
vif(mess_L_22_all_current)
#   Опыт собрания
vif(meet_L_01_all_exp)
vif(meet_L_07_all_exp)
vif(meet_L_22_all_exp)
#   Опыт рассылка
vif(mess_L_01_all_exp)
vif(mess_L_07_all_exp)
vif(mess_L_22_all_exp)
# Вывод автокорреляции нет

# 6. Проверка на эндогенность, если остатки коррелируют с чем-нибудь, все плохо
#   Текущая связь собрания

independent_vars_meet_L_01_all_current <- data_meet[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_L_01_all_current$residuals <- residuals_meet_L_01_all_current
cor_matrix_endog_meet_L_01_all_current <- cor(independent_vars_meet_L_01_all_current)
print(cor_matrix_endog_meet_L_01_all_current)
corrplot(cor_matrix_endog_meet_L_01_all_current)

independent_vars_meet_L_07_all_current <- data_meet[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_L_07_all_current$residuals <- residuals_meet_L_07_all_current
cor_matrix_endog_meet_L_07_all_current <- cor(independent_vars_meet_L_07_all_current)
print(cor_matrix_endog_meet_L_07_all_current)
corrplot(cor_matrix_endog_meet_L_07_all_current)

independent_vars_meet_L_22_all_current <- data_meet[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_L_22_all_current$residuals <- residuals_meet_L_22_all_current
cor_matrix_endog_meet_L_22_all_current <- cor(independent_vars_meet_L_22_all_current)
print(cor_matrix_endog_meet_L_22_all_current)
corrplot(cor_matrix_endog_meet_L_22_all_current)

#   Текущая связь рассылка

independent_vars_mess_L_01_all_current <- data_mess[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_L_01_all_current$residuals <- residuals_mess_L_01_all_current
cor_matrix_endog_mess_L_01_all_current <- cor(independent_vars_mess_L_01_all_current)
print(cor_matrix_endog_mess_L_01_all_current)
corrplot(cor_matrix_endog_mess_L_01_all_current)

independent_vars_mess_L_07_all_current <- data_mess[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_L_07_all_current$residuals <- residuals_mess_L_07_all_current
cor_matrix_endog_mess_L_07_all_current <- cor(independent_vars_mess_L_07_all_current)
print(cor_matrix_endog_mess_L_07_all_current)
corrplot(cor_matrix_endog_mess_L_07_all_current)

independent_vars_mess_L_22_all_current <- data_mess[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_L_22_all_current$residuals <- residuals_mess_L_22_all_current
cor_matrix_endog_mess_L_22_all_current <- cor(independent_vars_mess_L_22_all_current)
print(cor_matrix_endog_mess_L_22_all_current)
corrplot(cor_matrix_endog_mess_L_22_all_current)

#   Опыт собрания

independent_vars_meet_L_01_all_exp <- data_meet[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_L_01_all_exp$residuals <- residuals_meet_L_01_all_exp
cor_matrix_endog_meet_L_01_all_exp <- cor(independent_vars_meet_L_01_all_exp)
print(cor_matrix_endog_meet_L_01_all_exp)
corrplot(cor_matrix_endog_meet_L_01_all_exp)

independent_vars_meet_L_07_all_exp <- data_meet[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_L_07_all_exp$residuals <- residuals_meet_L_07_all_exp
cor_matrix_endog_meet_L_07_all_exp <- cor(independent_vars_meet_L_07_all_exp)
print(cor_matrix_endog_meet_L_07_all_exp)
corrplot(cor_matrix_endog_meet_L_07_all_exp)

independent_vars_meet_L_22_all_exp <- data_meet[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_L_22_all_exp$residuals <- residuals_meet_L_22_all_exp
cor_matrix_endog_meet_L_22_all_exp <- cor(independent_vars_meet_L_22_all_exp)
print(cor_matrix_endog_meet_L_22_all_exp)
corrplot(cor_matrix_endog_meet_L_22_all_exp)

#   Опыт рассылка

independent_vars_mess_L_01_all_exp <- data_mess[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_L_01_all_exp$residuals <- residuals_mess_L_01_all_exp
cor_matrix_endog_mess_L_01_all_exp <- cor(independent_vars_mess_L_01_all_exp)
print(cor_matrix_endog_mess_L_01_all_exp)
corrplot(cor_matrix_endog_mess_L_01_all_exp)

independent_vars_mess_L_07_all_exp <- data_mess[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_L_07_all_exp$residuals <- residuals_mess_L_07_all_exp
cor_matrix_endog_mess_L_07_all_exp <- cor(independent_vars_mess_L_07_all_exp)
print(cor_matrix_endog_mess_L_07_all_exp)
corrplot(cor_matrix_endog_mess_L_07_all_exp)

independent_vars_mess_L_22_all_exp <- data_mess[, c("ln_act", "lev", "ros", "gov_prop", "div", "share_cap_diff", "mer_acq", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_L_22_all_exp$residuals <- residuals_mess_L_22_all_exp
cor_matrix_endog_mess_L_22_all_exp <- cor(independent_vars_mess_L_22_all_exp)
print(cor_matrix_endog_mess_L_22_all_exp)
corrplot(cor_matrix_endog_mess_L_22_all_exp)

# Вывод остатки не коррелируют с переменными

# !!!Модели до 14 и после 14!!!


# Решения ОКНО от 0 до +7

meet_before2014L_07_all_current <- lm(data = data_meet_before2014, CAR_L_07_meet ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_before2014L_07_all_current)
meet_after2014L_07_all_current <- lm(data = data_meet_after2014, CAR_L_07_meet ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_after2014L_07_all_current)
meet_before2014L_07_all_exp <- lm(data = data_meet_before2014, CAR_L_07_meet ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_before2014L_07_all_exp)
meet_after2014L_07_all_exp <- lm(data = data_meet_after2014, CAR_L_07_meet ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_after2014L_07_all_exp)

#!!!#
stargazer(meet_before2014L_07_all_current, meet_after2014L_07_all_current, type = "text")
stargazer(meet_before2014L_07_all_exp, meet_after2014L_07_all_exp, type = "text")
#!!!#

# Рассылка ОКНО от 0 до +7

mess_before2014L_07_all_current <- lm(data = data_mess_before2014, CAR_L_07_mess ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_before2014L_07_all_current)
mess_after2014L_07_all_current <- lm(data = data_mess_after2014, CAR_L_07_mess ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_after2014L_07_all_current)
mess_before2014L_07_all_exp <- lm(data = data_mess_before2014, CAR_L_07_mess ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_before2014L_07_all_exp)
mess_after2014L_07_all_exp <- lm(data = data_mess_after2014, CAR_L_07_mess ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_after2014L_07_all_exp)

#!!!#
stargazer(mess_before2014L_07_all_current, mess_after2014L_07_all_current, type = "text")
stargazer(mess_before2014L_07_all_exp, mess_after2014L_07_all_exp, type = "text")
#!!!#

# Решения ОКНО от 0 до +1

meet_before2014L_01_all_current <- lm(data = data_meet_before2014, CAR_L_01_meet ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_before2014L_01_all_current)
meet_after2014L_01_all_current <- lm(data = data_meet_after2014, CAR_L_01_meet ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_after2014L_01_all_current)
meet_before2014L_01_all_exp <- lm(data = data_meet_before2014, CAR_L_01_meet ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_before2014L_01_all_exp)
meet_after2014L_01_all_exp <- lm(data = data_meet_after2014, CAR_L_01_meet ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_after2014L_01_all_exp)

#!!!#
stargazer(meet_before2014L_01_all_current, meet_after2014L_01_all_current, type = "text")
stargazer(meet_before2014L_01_all_exp, meet_after2014L_01_all_exp, type = "text")
#!!!#

# Рассылка ОКНО от 0 до +1

mess_before2014L_01_all_current <- lm(data = data_mess_before2014, CAR_L_01_mess ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_before2014L_01_all_current)
mess_after2014L_01_all_current <- lm(data = data_mess_after2014, CAR_L_01_mess ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_after2014L_01_all_current)
mess_before2014L_01_all_exp <- lm(data = data_mess_before2014, CAR_L_01_mess ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_before2014L_01_all_exp)
mess_after2014L_01_all_exp <- lm(data = data_mess_after2014, CAR_L_01_mess ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_after2014L_01_all_exp)

#!!!#
stargazer(mess_before2014L_01_all_current, mess_after2014L_01_all_current, type = "text")
stargazer(mess_before2014L_01_all_exp, mess_after2014L_01_all_exp, type = "text")
#!!!#

# Решения ОКНО от -2 до +2

meet_before2014L_22_all_current <- lm(data = data_meet_before2014, CAR_L_22_meet ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_before2014L_22_all_current)
meet_after2014L_22_all_current <- lm(data = data_meet_after2014, CAR_L_22_meet ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(meet_after2014L_22_all_current)
meet_before2014L_22_all_exp <- lm(data = data_meet_before2014, CAR_L_22_meet ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_before2014L_22_all_exp)
meet_after2014L_22_all_exp <- lm(data = data_meet_after2014, CAR_L_22_meet ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(meet_after2014L_22_all_exp)

#!!!#
stargazer(meet_before2014L_22_all_current, meet_after2014L_22_all_current, type = "text")
stargazer(meet_before2014L_22_all_exp, meet_after2014L_22_all_exp, type = "text")
#!!!#

# Рассылка ОКНО от -2 до +2

mess_before2014L_22_all_current <- lm(data = data_mess_before2014, CAR_L_22_mess ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_before2014L_22_all_current)
mess_after2014L_22_all_current <- lm(data = data_mess_after2014, CAR_L_22_mess ~ gov_prop + pre_board + diff_gen + strong_current + weak_current + strong_gen + weak_gen + polit_gen)
summary(mess_after2014L_22_all_current)
mess_before2014L_22_all_exp <- lm(data = data_mess_before2014, CAR_L_22_mess ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_before2014L_22_all_exp)
mess_after2014L_22_all_exp <- lm(data = data_mess_after2014, CAR_L_22_mess ~ gov_prop + pre_board + diff_gen + strong_exp + weak_exp + strong_gen + weak_gen + polit_gen)
summary(mess_after2014L_22_all_exp)

#!!!#
stargazer(mess_before2014L_22_all_current, mess_after2014L_22_all_current, type = "text")
stargazer(mess_before2014L_22_all_exp, mess_after2014L_22_all_exp, type = "text")
#!!!#


# 1. Проверка спецификации (квадраты и кубы) resettest нулевая гипотеза все правильно, т.е если p-value>0.05, то не отклоняем, все ок
#   Текущая связь собрания
resettest(meet_before2014L_01_all_current, power = 2:3, type = "regressor") 
resettest(meet_before2014L_07_all_current, power = 2:3, type = "regressor")
resettest(meet_before2014L_22_all_current, power = 2:3, type = "regressor")
resettest(meet_after2014L_01_all_current, power = 2:3, type = "regressor") 
resettest(meet_after2014L_07_all_current, power = 2:3, type = "regressor")
resettest(meet_after2014L_22_all_current, power = 2:3, type = "regressor")
#   Текущая связь рассылка
resettest(mess_before2014L_01_all_current, power = 2:3, type = "regressor") 
resettest(mess_before2014L_07_all_current, power = 2:3, type = "regressor")
resettest(mess_before2014L_22_all_current, power = 2:3, type = "regressor")
resettest(mess_after2014L_01_all_current, power = 2:3, type = "regressor") 
resettest(mess_after2014L_07_all_current, power = 2:3, type = "regressor")
resettest(mess_after2014L_22_all_current, power = 2:3, type = "regressor")
#   Опыт собрания
resettest(meet_before2014L_01_all_exp, power = 2:3, type = "regressor") 
resettest(meet_before2014L_07_all_exp, power = 2:3, type = "regressor")
resettest(meet_before2014L_22_all_exp, power = 2:3, type = "regressor")
resettest(meet_after2014L_01_all_exp, power = 2:3, type = "regressor") 
resettest(meet_after2014L_07_all_exp, power = 2:3, type = "regressor")
resettest(meet_after2014L_22_all_exp, power = 2:3, type = "regressor")
#   Опыт рассылка
resettest(mess_before2014L_01_all_exp, power = 2:3, type = "regressor") 
resettest(mess_before2014L_07_all_exp, power = 2:3, type = "regressor")
resettest(mess_before2014L_22_all_exp, power = 2:3, type = "regressor")
resettest(mess_after2014L_01_all_exp, power = 2:3, type = "regressor") 
resettest(mess_after2014L_07_all_exp, power = 2:3, type = "regressor")
resettest(mess_after2014L_22_all_exp, power = 2:3, type = "regressor")
# Вывод, квадраты и кубы не нужны нигде


# 2. Равенство математических ожиданий ошибок нулю (иначе смещенный результат)
#   Нулевая гипотеза что ошибки 0, если p-value > 0.05 не отклоняем, все ок
#   Текущая связь собрания
residuals_meet_before2014L_01_all_current <- resid(meet_before2014L_01_all_current) 
t_test_meet_before2014L_01_all_current <- t.test(residuals_meet_before2014L_01_all_current, mu = 0)
print(t_test_meet_before2014L_01_all_current)
residuals_meet_before2014L_07_all_current <- resid(meet_before2014L_07_all_current) 
t_test_meet_before2014L_07_all_current <- t.test(residuals_meet_before2014L_07_all_current, mu = 0)
print(t_test_meet_before2014L_07_all_current)
residuals_meet_before2014L_22_all_current <- resid(meet_before2014L_22_all_current) 
t_test_meet_before2014L_22_all_current <- t.test(residuals_meet_before2014L_22_all_current, mu = 0)
print(t_test_meet_before2014L_22_all_current)
residuals_meet_after2014L_01_all_current <- resid(meet_after2014L_01_all_current) 
t_test_meet_after2014L_01_all_current <- t.test(residuals_meet_after2014L_01_all_current, mu = 0)
print(t_test_meet_after2014L_01_all_current)
residuals_meet_after2014L_07_all_current <- resid(meet_after2014L_07_all_current) 
t_test_meet_after2014L_07_all_current <- t.test(residuals_meet_after2014L_07_all_current, mu = 0)
print(t_test_meet_after2014L_07_all_current)
residuals_meet_after2014L_22_all_current <- resid(meet_after2014L_22_all_current) 
t_test_meet_after2014L_22_all_current <- t.test(residuals_meet_after2014L_22_all_current, mu = 0)
print(t_test_meet_after2014L_22_all_current)
#   Текущая связь рассылка
residuals_mess_before2014L_01_all_current <- resid(mess_before2014L_01_all_current) 
t_test_mess_before2014L_01_all_current <- t.test(residuals_mess_before2014L_01_all_current, mu = 0)
print(t_test_mess_before2014L_01_all_current)
residuals_mess_before2014L_07_all_current <- resid(mess_before2014L_07_all_current) 
t_test_mess_before2014L_07_all_current <- t.test(residuals_mess_before2014L_07_all_current, mu = 0)
print(t_test_mess_before2014L_07_all_current)
residuals_mess_before2014L_22_all_current <- resid(mess_before2014L_22_all_current) 
t_test_mess_before2014L_22_all_current <- t.test(residuals_mess_before2014L_22_all_current, mu = 0)
print(t_test_mess_before2014L_22_all_current)
residuals_mess_after2014L_01_all_current <- resid(mess_after2014L_01_all_current) 
t_test_mess_after2014L_01_all_current <- t.test(residuals_mess_after2014L_01_all_current, mu = 0)
print(t_test_mess_after2014L_01_all_current)
residuals_mess_after2014L_07_all_current <- resid(mess_after2014L_07_all_current) 
t_test_mess_after2014L_07_all_current <- t.test(residuals_mess_after2014L_07_all_current, mu = 0)
print(t_test_mess_after2014L_07_all_current)
residuals_mess_after2014L_22_all_current <- resid(mess_after2014L_22_all_current) 
t_test_mess_after2014L_22_all_current <- t.test(residuals_mess_after2014L_22_all_current, mu = 0)
print(t_test_mess_after2014L_22_all_current)
#   Опыт собрания
residuals_meet_before2014L_01_all_exp <- resid(meet_before2014L_01_all_exp) 
t_test_meet_before2014L_01_all_exp <- t.test(residuals_meet_before2014L_01_all_exp, mu = 0)
print(t_test_meet_before2014L_01_all_exp)
residuals_meet_before2014L_07_all_exp <- resid(meet_before2014L_07_all_exp) 
t_test_meet_before2014L_07_all_exp <- t.test(residuals_meet_before2014L_07_all_exp, mu = 0)
print(t_test_meet_before2014L_07_all_exp)
residuals_meet_before2014L_22_all_exp <- resid(meet_before2014L_22_all_exp) 
t_test_meet_before2014L_22_all_exp <- t.test(residuals_meet_before2014L_22_all_exp, mu = 0)
print(t_test_meet_before2014L_22_all_exp)
residuals_meet_after2014L_01_all_exp <- resid(meet_after2014L_01_all_exp) 
t_test_meet_after2014L_01_all_exp <- t.test(residuals_meet_after2014L_01_all_exp, mu = 0)
print(t_test_meet_after2014L_01_all_exp)
residuals_meet_after2014L_07_all_exp <- resid(meet_after2014L_07_all_exp) 
t_test_meet_after2014L_07_all_exp <- t.test(residuals_meet_after2014L_07_all_exp, mu = 0)
print(t_test_meet_after2014L_07_all_exp)
residuals_meet_after2014L_22_all_exp <- resid(meet_after2014L_22_all_exp) 
t_test_meet_after2014L_22_all_exp <- t.test(residuals_meet_after2014L_22_all_exp, mu = 0)
print(t_test_meet_after2014L_22_all_exp)
#   Опыт рассылка
residuals_mess_before2014L_01_all_exp <- resid(mess_before2014L_01_all_exp) 
t_test_mess_before2014L_01_all_exp <- t.test(residuals_mess_before2014L_01_all_exp, mu = 0)
print(t_test_mess_before2014L_01_all_exp)
residuals_mess_before2014L_07_all_exp <- resid(mess_before2014L_07_all_exp) 
t_test_mess_before2014L_07_all_exp <- t.test(residuals_mess_before2014L_07_all_exp, mu = 0)
print(t_test_mess_before2014L_07_all_exp)
residuals_mess_before2014L_22_all_exp <- resid(mess_before2014L_22_all_exp) 
t_test_mess_before2014L_22_all_exp <- t.test(residuals_mess_before2014L_22_all_exp, mu = 0)
print(t_test_mess_before2014L_22_all_exp)
residuals_mess_after2014L_01_all_exp <- resid(mess_after2014L_01_all_exp) 
t_test_mess_after2014L_01_all_exp <- t.test(residuals_mess_after2014L_01_all_exp, mu = 0)
print(t_test_mess_after2014L_01_all_exp)
residuals_mess_after2014L_07_all_exp <- resid(mess_after2014L_07_all_exp) 
t_test_mess_after2014L_07_all_exp <- t.test(residuals_mess_after2014L_07_all_exp, mu = 0)
print(t_test_mess_after2014L_07_all_exp)
residuals_mess_after2014L_22_all_exp <- resid(mess_after2014L_22_all_exp) 
t_test_mess_after2014L_22_all_exp <- t.test(residuals_mess_after2014L_22_all_exp, mu = 0)
print(t_test_mess_after2014L_22_all_exp)
# Вывод все ок, ошибки не отличаются от нуля, результат несмещенный

# 3. Проверка гомоскедастичности (одинаковая дисперсия остатков)
# bptest нулевая гипотеза все хорошо гомоскедастичность, если p-value > 0.05 не отклоняем, все ок
#   Текущая связь собрания
bptest(meet_before2014L_01_all_current)
bptest(meet_before2014L_07_all_current)
bptest(meet_before2014L_22_all_current)
bptest(meet_after2014L_01_all_current)
bptest(meet_after2014L_07_all_current)
bptest(meet_after2014L_22_all_current)
#   Текущая связь рассылка
bptest(mess_before2014L_01_all_current)
bptest(mess_before2014L_07_all_current)
bptest(mess_before2014L_22_all_current)
bptest(mess_after2014L_01_all_current)
bptest(mess_after2014L_07_all_current)
bptest(mess_after2014L_22_all_current)
#   Опыт собрания
bptest(meet_before2014L_01_all_exp)
bptest(meet_before2014L_07_all_exp)
bptest(meet_before2014L_22_all_exp)
bptest(meet_after2014L_01_all_exp)
bptest(meet_after2014L_07_all_exp)
bptest(meet_after2014L_22_all_exp)
#   Опыт рассылка
bptest(mess_before2014L_01_all_exp)
bptest(mess_before2014L_07_all_exp)
bptest(mess_before2014L_22_all_exp)
bptest(mess_after2014L_01_all_exp)
bptest(mess_after2014L_07_all_exp)
bptest(mess_after2014L_22_all_exp)
# Вывод почти все стало ок, кроме 01 после 14 meet all current, meet all exp

# 4.Проверка автокорреляции ошибок
# тест дарбина уотсона, нулевая гипотеза автокорреляции ПЕРВОГО ПОРЯДКА нет, если p-value > 0.05 все ок

#   Текущая связь собрания
dw_test_meet_before2014L_01_all_current <- dwtest(meet_before2014L_01_all_current)
print(dw_test_meet_before2014L_01_all_current)
dw_test_meet_before2014L_07_all_current <- dwtest(meet_before2014L_07_all_current)
print(dw_test_meet_before2014L_07_all_current)
dw_test_meet_before2014L_22_all_current <- dwtest(meet_before2014L_22_all_current)
print(dw_test_meet_before2014L_22_all_current)
dw_test_meet_after2014L_01_all_current <- dwtest(meet_after2014L_01_all_current)
print(dw_test_meet_after2014L_01_all_current)
dw_test_meet_after2014L_07_all_current <- dwtest(meet_after2014L_07_all_current)
print(dw_test_meet_after2014L_07_all_current)
dw_test_meet_after2014L_22_all_current <- dwtest(meet_after2014L_22_all_current)
print(dw_test_meet_after2014L_22_all_current)
#   Текущая связь рассылка
dw_test_mess_before2014L_01_all_current <- dwtest(mess_before2014L_01_all_current)
print(dw_test_mess_before2014L_01_all_current)
dw_test_mess_before2014L_07_all_current <- dwtest(mess_before2014L_07_all_current)
print(dw_test_mess_before2014L_07_all_current)
dw_test_mess_before2014L_22_all_current <- dwtest(mess_before2014L_22_all_current)
print(dw_test_mess_before2014L_22_all_current)
dw_test_mess_after2014L_01_all_current <- dwtest(mess_after2014L_01_all_current)
print(dw_test_mess_after2014L_01_all_current)
dw_test_mess_after2014L_07_all_current <- dwtest(mess_after2014L_07_all_current)
print(dw_test_mess_after2014L_07_all_current)
dw_test_mess_after2014L_22_all_current <- dwtest(mess_after2014L_22_all_current)
print(dw_test_mess_after2014L_22_all_current)
#   Опыт собрания
dw_test_meet_before2014L_01_all_exp <- dwtest(meet_before2014L_01_all_exp)
print(dw_test_meet_before2014L_01_all_exp)
dw_test_meet_before2014L_07_all_exp <- dwtest(meet_before2014L_07_all_exp)
print(dw_test_meet_before2014L_07_all_exp)
dw_test_meet_before2014L_22_all_exp <- dwtest(meet_before2014L_22_all_exp)
print(dw_test_meet_before2014L_22_all_exp)
dw_test_meet_after2014L_01_all_exp <- dwtest(meet_after2014L_01_all_exp)
print(dw_test_meet_after2014L_01_all_exp)
dw_test_meet_after2014L_07_all_exp <- dwtest(meet_after2014L_07_all_exp)
print(dw_test_meet_after2014L_07_all_exp)
dw_test_meet_after2014L_22_all_exp <- dwtest(meet_after2014L_22_all_exp)
print(dw_test_meet_after2014L_22_all_exp)
#   Опыт рассылка
dw_test_mess_before2014L_01_all_exp <- dwtest(mess_before2014L_01_all_exp)
print(dw_test_mess_before2014L_01_all_exp)
dw_test_mess_before2014L_07_all_exp <- dwtest(mess_before2014L_07_all_exp)
print(dw_test_mess_before2014L_07_all_exp)
dw_test_mess_before2014L_22_all_exp <- dwtest(mess_before2014L_22_all_exp)
print(dw_test_mess_before2014L_22_all_exp)
dw_test_mess_after2014L_01_all_exp <- dwtest(mess_after2014L_01_all_exp)
print(dw_test_mess_after2014L_01_all_exp)
dw_test_mess_after2014L_07_all_exp <- dwtest(mess_after2014L_07_all_exp)
print(dw_test_mess_after2014L_07_all_exp)
dw_test_mess_after2014L_22_all_exp <- dwtest(mess_after2014L_22_all_exp)
print(dw_test_mess_after2014L_22_all_exp)
# Вывод все ок, кроме после 14 (01), те в 4 моделях

# 5. Проверка мультиколлинеарности
# vif

#   Текущая связь собрания
vif(meet_before2014L_01_all_current)
vif(meet_before2014L_07_all_current)
vif(meet_before2014L_22_all_current)
vif(meet_after2014L_01_all_current)
vif(meet_after2014L_07_all_current)
vif(meet_after2014L_22_all_current)
#   Текущая связь рассылка
vif(mess_before2014L_01_all_current)
vif(mess_before2014L_07_all_current)
vif(mess_before2014L_22_all_current)
vif(mess_after2014L_01_all_current)
vif(mess_after2014L_07_all_current)
vif(mess_after2014L_22_all_current)
#   Опыт собрания
vif(meet_before2014L_01_all_exp)
vif(meet_before2014L_07_all_exp)
vif(meet_before2014L_22_all_exp)
vif(meet_after2014L_01_all_exp)
vif(meet_after2014L_07_all_exp)
vif(meet_after2014L_22_all_exp)
#   Опыт рассылка
vif(mess_before2014L_01_all_exp)
vif(mess_before2014L_07_all_exp)
vif(mess_before2014L_22_all_exp)
vif(mess_after2014L_01_all_exp)
vif(mess_after2014L_07_all_exp)
vif(mess_after2014L_22_all_exp)
# Вывод автокорреляции нет

# 6. Проверка на эндогенность, если остатки коррелируют с чем-нибудь, все плохо
#   Текущая связь собрания
independent_vars_meet_before2014L_01_all_current <- data_meet_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_before2014L_01_all_current$residuals <- residuals_meet_before2014L_01_all_current
cor_matrix_endog_meet_before2014L_01_all_current <- cor(independent_vars_meet_before2014L_01_all_current)
print(cor_matrix_endog_meet_before2014L_01_all_current)
corrplot(cor_matrix_endog_meet_before2014L_01_all_current)

independent_vars_meet_before2014L_07_all_current <- data_meet_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_before2014L_07_all_current$residuals <- residuals_meet_before2014L_07_all_current
cor_matrix_endog_meet_before2014L_07_all_current <- cor(independent_vars_meet_before2014L_07_all_current)
print(cor_matrix_endog_meet_before2014L_07_all_current)
corrplot(cor_matrix_endog_meet_before2014L_07_all_current)

independent_vars_meet_before2014L_22_all_current <- data_meet_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_before2014L_22_all_current$residuals <- residuals_meet_before2014L_22_all_current
cor_matrix_endog_meet_before2014L_22_all_current <- cor(independent_vars_meet_before2014L_22_all_current)
print(cor_matrix_endog_meet_before2014L_22_all_current)
corrplot(cor_matrix_endog_meet_before2014L_22_all_current)

independent_vars_meet_after2014L_01_all_current <- data_meet_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_after2014L_01_all_current$residuals <- residuals_meet_after2014L_01_all_current
cor_matrix_endog_meet_after2014L_01_all_current <- cor(independent_vars_meet_after2014L_01_all_current)
print(cor_matrix_endog_meet_after2014L_01_all_current)
corrplot(cor_matrix_endog_meet_after2014L_01_all_current)

independent_vars_meet_after2014L_07_all_current <- data_meet_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_after2014L_07_all_current$residuals <- residuals_meet_after2014L_07_all_current
cor_matrix_endog_meet_after2014L_07_all_current <- cor(independent_vars_meet_after2014L_07_all_current)
print(cor_matrix_endog_meet_after2014L_07_all_current)
corrplot(cor_matrix_endog_meet_after2014L_07_all_current)

independent_vars_meet_after2014L_22_all_current <- data_meet_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_after2014L_22_all_current$residuals <- residuals_meet_after2014L_22_all_current
cor_matrix_endog_meet_after2014L_22_all_current <- cor(independent_vars_meet_after2014L_22_all_current)
print(cor_matrix_endog_meet_after2014L_22_all_current)
corrplot(cor_matrix_endog_meet_after2014L_22_all_current)

#   Текущая связь рассылка
independent_vars_mess_before2014L_01_all_current <- data_mess_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_before2014L_01_all_current$residuals <- residuals_mess_before2014L_01_all_current
cor_matrix_endog_mess_before2014L_01_all_current <- cor(independent_vars_mess_before2014L_01_all_current)
print(cor_matrix_endog_mess_before2014L_01_all_current)
corrplot(cor_matrix_endog_mess_before2014L_01_all_current)

independent_vars_mess_before2014L_07_all_current <- data_mess_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_before2014L_07_all_current$residuals <- residuals_mess_before2014L_07_all_current
cor_matrix_endog_mess_before2014L_07_all_current <- cor(independent_vars_mess_before2014L_07_all_current)
print(cor_matrix_endog_mess_before2014L_07_all_current)
corrplot(cor_matrix_endog_mess_before2014L_07_all_current)

independent_vars_mess_before2014L_22_all_current <- data_mess_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_before2014L_22_all_current$residuals <- residuals_mess_before2014L_22_all_current
cor_matrix_endog_mess_before2014L_22_all_current <- cor(independent_vars_mess_before2014L_22_all_current)
print(cor_matrix_endog_mess_before2014L_22_all_current)
corrplot(cor_matrix_endog_mess_before2014L_22_all_current)

independent_vars_mess_after2014L_01_all_current <- data_mess_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_after2014L_01_all_current$residuals <- residuals_mess_after2014L_01_all_current
cor_matrix_endog_mess_after2014L_01_all_current <- cor(independent_vars_mess_after2014L_01_all_current)
print(cor_matrix_endog_mess_after2014L_01_all_current)
corrplot(cor_matrix_endog_mess_after2014L_01_all_current)

independent_vars_mess_after2014L_07_all_current <- data_mess_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_after2014L_07_all_current$residuals <- residuals_mess_after2014L_07_all_current
cor_matrix_endog_mess_after2014L_07_all_current <- cor(independent_vars_mess_after2014L_07_all_current)
print(cor_matrix_endog_mess_after2014L_07_all_current)
corrplot(cor_matrix_endog_mess_after2014L_07_all_current)

independent_vars_mess_after2014L_22_all_current <- data_mess_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_current", "weak_current", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_after2014L_22_all_current$residuals <- residuals_mess_after2014L_22_all_current
cor_matrix_endog_mess_after2014L_22_all_current <- cor(independent_vars_mess_after2014L_22_all_current)
print(cor_matrix_endog_mess_after2014L_22_all_current)
corrplot(cor_matrix_endog_mess_after2014L_22_all_current)

#   Опыт собрания

independent_vars_meet_before2014L_01_all_exp <- data_meet_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_before2014L_01_all_exp$residuals <- residuals_meet_before2014L_01_all_exp
cor_matrix_endog_meet_before2014L_01_all_exp <- cor(independent_vars_meet_before2014L_01_all_exp)
print(cor_matrix_endog_meet_before2014L_01_all_exp)
corrplot(cor_matrix_endog_meet_before2014L_01_all_exp)

independent_vars_meet_before2014L_07_all_exp <- data_meet_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_before2014L_07_all_exp$residuals <- residuals_meet_before2014L_07_all_exp
cor_matrix_endog_meet_before2014L_07_all_exp <- cor(independent_vars_meet_before2014L_07_all_exp)
print(cor_matrix_endog_meet_before2014L_07_all_exp)
corrplot(cor_matrix_endog_meet_before2014L_07_all_exp)

independent_vars_meet_before2014L_22_all_exp <- data_meet_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_before2014L_22_all_exp$residuals <- residuals_meet_before2014L_22_all_exp
cor_matrix_endog_meet_before2014L_22_all_exp <- cor(independent_vars_meet_before2014L_22_all_exp)
print(cor_matrix_endog_meet_before2014L_22_all_exp)
corrplot(cor_matrix_endog_meet_before2014L_22_all_exp)

independent_vars_meet_after2014L_01_all_exp <- data_meet_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_after2014L_01_all_exp$residuals <- residuals_meet_after2014L_01_all_exp
cor_matrix_endog_meet_after2014L_01_all_exp <- cor(independent_vars_meet_after2014L_01_all_exp)
print(cor_matrix_endog_meet_after2014L_01_all_exp)
corrplot(cor_matrix_endog_meet_after2014L_01_all_exp)

independent_vars_meet_after2014L_07_all_exp <- data_meet_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_after2014L_07_all_exp$residuals <- residuals_meet_after2014L_07_all_exp
cor_matrix_endog_meet_after2014L_07_all_exp <- cor(independent_vars_meet_after2014L_07_all_exp)
print(cor_matrix_endog_meet_after2014L_07_all_exp)
corrplot(cor_matrix_endog_meet_after2014L_07_all_exp)

independent_vars_meet_after2014L_22_all_exp <- data_meet_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_meet_after2014L_22_all_exp$residuals <- residuals_meet_after2014L_22_all_exp
cor_matrix_endog_meet_after2014L_22_all_exp <- cor(independent_vars_meet_after2014L_22_all_exp)
print(cor_matrix_endog_meet_after2014L_22_all_exp)
corrplot(cor_matrix_endog_meet_after2014L_22_all_exp)

#   Опыт рассылка

independent_vars_mess_before2014L_01_all_exp <- data_mess_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_before2014L_01_all_exp$residuals <- residuals_mess_before2014L_01_all_exp
cor_matrix_endog_mess_before2014L_01_all_exp <- cor(independent_vars_mess_before2014L_01_all_exp)
print(cor_matrix_endog_mess_before2014L_01_all_exp)
corrplot(cor_matrix_endog_mess_before2014L_01_all_exp)

independent_vars_mess_before2014L_07_all_exp <- data_mess_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_before2014L_07_all_exp$residuals <- residuals_mess_before2014L_07_all_exp
cor_matrix_endog_mess_before2014L_07_all_exp <- cor(independent_vars_mess_before2014L_07_all_exp)
print(cor_matrix_endog_mess_before2014L_07_all_exp)
corrplot(cor_matrix_endog_mess_before2014L_07_all_exp)

independent_vars_mess_before2014L_22_all_exp <- data_mess_before2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_before2014L_22_all_exp$residuals <- residuals_mess_before2014L_22_all_exp
cor_matrix_endog_mess_before2014L_22_all_exp <- cor(independent_vars_mess_before2014L_22_all_exp)
print(cor_matrix_endog_mess_before2014L_22_all_exp)
corrplot(cor_matrix_endog_mess_before2014L_22_all_exp)

independent_vars_mess_after2014L_01_all_exp <- data_mess_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_after2014L_01_all_exp$residuals <- residuals_mess_after2014L_01_all_exp
cor_matrix_endog_mess_after2014L_01_all_exp <- cor(independent_vars_mess_after2014L_01_all_exp)
print(cor_matrix_endog_mess_after2014L_01_all_exp)
corrplot(cor_matrix_endog_mess_after2014L_01_all_exp)

independent_vars_mess_after2014L_07_all_exp <- data_mess_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_after2014L_07_all_exp$residuals <- residuals_mess_after2014L_07_all_exp
cor_matrix_endog_mess_after2014L_07_all_exp <- cor(independent_vars_mess_after2014L_07_all_exp)
print(cor_matrix_endog_mess_after2014L_07_all_exp)
corrplot(cor_matrix_endog_mess_after2014L_07_all_exp)

independent_vars_mess_after2014L_22_all_exp <- data_mess_after2014[, c("gov_prop", "pre_board", "diff_gen", "strong_exp", "weak_exp", "strong_gen", "weak_gen", "polit_gen")]
independent_vars_mess_after2014L_22_all_exp$residuals <- residuals_mess_after2014L_22_all_exp
cor_matrix_endog_mess_after2014L_22_all_exp <- cor(independent_vars_mess_after2014L_22_all_exp)
print(cor_matrix_endog_mess_after2014L_22_all_exp)
corrplot(cor_matrix_endog_mess_after2014L_22_all_exp)

# Вывод остатки не коррелируют с переменными

summary(data_meet)
# AIC и BIC mtable(m_20_5_wk, m_20_5_wk_nonlinear, summary.stats=c("R-squared", "adj. R-squared","F","p","N", "AIC","BIC"))

