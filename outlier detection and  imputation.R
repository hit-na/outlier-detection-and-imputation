#--------------------#
# データクレンジング #
# (外れ値検出と補完) #
#--------------------#

## 説明＆手順＆注意事項 ##
# ●説明
# 対象データの外れ値を検出し欠損値に置き換え線形補完するコード
# 外れ値検出・補完結果のcsvファイルとhtmlグラフを出力します。
# 前処理：対数化してTukey's fences(k=±3)で大きな外れ値と0以下の値を欠損値の置き換え
# 本処理：3周期(1,24,168)のSARIMAモデルから共通するAOを検出し欠損値に置き換え
# 補完  ：欠損値を線形補間
# ●手順
# Ctrl + A で全選択して Ctrl + Enter で実行
# 列選択は「日時列」と「処理するデータ列」を選択
# ●注意事項
# 外れ値検出を行う対象ファイルはcsvファイル
# csvファイルは、1行目に列名、2行目以降に観測値で、1列目は日時のフォーマット
# ファイル選択と列選択、日付入力画面が背面に隠れることがある
# 読み込むcsvファイルのフォーマット:1行目が列名、2行目以降に観測値
# 日時の対応フォーマット:"%Y-%m-%d %H:%M:%S" , "%Y/%m/%d %H:%M"


#### パッケージのインストールとライブラリの読み込み ####
# 必要なパッケージのリスト
required_packages <- c("readr", "tcltk","tidyverse","forecast", "tsoutliers",
                       "plotly","htmlwidgets")

# パッケージがインストールされているか確認し、インストールされていなければインストール
install_if_missing <- function(p) {
  if (!requireNamespace(p, quietly = TRUE)) {
    install.packages(p, dependencies = TRUE)
  }
}

# 必要なパッケージをインストール
lapply(required_packages, install_if_missing)

# ライブラリを読み込み
lapply(required_packages, library, character.only = TRUE)





#### 関数の作成 ####
# ダイアログボックスで日付を入力する関数を作成
get_date_input <- function(prompt_text) {
  # ウィンドウの作成
  window <- tktoplevel()
  tkwm.title(window, prompt_text)
  
  # エントリーウィジェットの作成
  entry_var <- tclVar("")  # 変数を設定
  entry <- tkentry(window, textvariable = entry_var)  # 変数をエントリーにリンク
  tkgrid(tklabel(window, text = prompt_text))
  tkgrid(entry)
  
  # OKボタンの作成
  ok_button <- tkbutton(window, text = "OK", command = function() tkdestroy(window))
  tkgrid(ok_button)
  
  # ウィンドウを待機
  tkwait.window(window)
  
  # 入力された値を取得
  date_input <- tclvalue(entry_var)
  return(date_input)
}


# 前処理：対象データを対数化してTukey's fences
# 論文2章冒頭_式(1)：
# 下限 = Q1 - k * (Q3 - Q1), 上限 = Q3 + ku * (Q3 - Q1)
# (本コードでは下限側係数 k, 上限係数 ku = kk として実装) [式(1)参照]
repOutliersNA <- function(Column,k,kk) {
  Column <- as_tibble(Column)
  # 0以下は対数化できないため1に補正
  Column[Column <= 0] <- 1
  lcolumn <- log10(Column)
  # 
  xx1 <- Column %>% plyr::mutate(
    Norm = (lcolumn - min(lcolumn, na.rm = T)) / (max(lcolumn, na.rm = T) - min(lcolumn, na.rm = T)) * (1-0) + 0
  )
  
  # 四分位数 Q1, Q3
  qq <- data.frame(quantile(xx1[[2]],c(0.25, 0.75), na.rm = T))
  Q1 <- qq[1, 1]
  Q3 <- qq[2, 1]
  
  # 式(1)の下限・上限フェンス
  outer_l_Q1 <- Q1 - k * (Q3 - Q1)
  outer_m_Q3 <- Q3 + kk * (Q3 - Q1)
  
  # フェンス外の行を抽出
  outer_ll <- which(xx1$Norm$value < outer_l_Q1)
  outer_mm <- which(xx1$Norm$value > outer_m_Q3)
  row_num_out <- unique(c(outer_ll, outer_mm)) %>% sort()
  
  # 外れ値をNAに
  outer_outlier <- cbind.data.frame(lcolumn[row_num_out,], row_number = row_num_out)
  Column_removeOutliers <- Column
    Column_removeOutliers[outer_outlier$row_number, 1] <- NA 
  
  # 元と同じ長さのベクトルを返す
  return(Column_removeOutliers$value)
}






#### 1.対象ファイルの選択と期間抽出####
# ファイル選択ダイアログを表示
file_name <- file.choose()

# CSVファイルを読み込む
data_ori <- read_csv(file_name, locale = locale(encoding = "SHIFT-JIS"))

# 利用可能な列名を表示して選択(日時と外れ値検出データの選択)
selected_columns <- tk_select.list(names(data_ori), 
                                   title = "Select columns to include:", 
                                   multiple = TRUE)

# 選択された列のみを抽出
data <- data_ori[, selected_columns]

# 列名を設定
names(data)[1] = "date" 
names(data)[2] = "value" 

# 日時フォーマットに基づいて変換
if (grepl("-", data$date[1])) {
  # "-" が含まれている場合_大学建物群（例: "2013-04-01 00:00:00"）
  data$date <- as.POSIXct(data$date, format = "%Y-%m-%d %H:%M:%S")
} else if (grepl("/", data$date[1])) {
  # "/" が含まれている場合_研修施設（例: "2023/04/01 00:00"）
  data$date <- as.POSIXct(data$date, format = "%Y/%m/%d %H:%M")
  # 2列目がcharacterの場合数値に変換
  data[ ,2:ncol(data)] <- apply(data[,2:ncol(data)],2,as.numeric)
} else {
  # その他のフォーマットの場合
  cat("日時フォーマットが認識できません。最初の値:", data$date, "\n")
}

# 開始日と終了日を入力（ダイアログ表示）
start_date_input <- get_date_input("開始日時を入力してください（例: 2013-04-01）")
end_date_input <- get_date_input("終了日時を入力してください（例: 2014-03-31）")

# 入力された日付文字列をPOSIXctに変換
start_date <- as.POSIXct(paste(start_date_input, "00:00:00"),format = "%Y-%m-%d %H:%M:%S")
end_date <- as.POSIXct(paste(end_date_input, "23:00:00"), format = "%Y-%m-%d %H:%M:%S")

# 指定された日付範囲のデータのみを抽出
data <- data[data$date >= start_date & data$date <= end_date, ]

# #時刻列のNA行を削除（NA行できることがあるため）
data <- data[!is.na(data$date),] 





#### 2.前処理:対数化 + Tukey’s fencesで極端な外れ値を除外(論文2章冒頭) ####
data$TF_val <- NA
# repOutliersNA(Tukey's fencesを行う列,下限のkの値,上限のkの値)
data$TF_val <- repOutliersNA(data$value, 3, 3)



#### 3.本処理:3周期SARIMAモデルでAO影響度推定 ####
# 論文2.1節
# 周期 s ∈ {1, 24, 168} のSARIMAモデルを作成する準備
datats1　<-　ts(data$TF_val,frequency = 1) 
datats24　<-　ts(data$TF_val,frequency = 24) 
datats168　<-　ts(data$TF_val,frequency = 168) 

# 周期別のSARIMAモデルを推定
fit_1 <- auto.arima(datats1,seasonal = TRUE,trace = TRUE) 
fit_24 <- auto.arima(datats24,seasonal = TRUE,trace = TRUE)
fit_168 <- auto.arima(datats168,seasonal = TRUE,trace = TRUE)

# 推定したモデルの予測値
fittedval_1 <- fitted(fit_1)
fittedval_24 <- fitted(fit_24)
fittedval_168 <- fitted(fit_168)

# 実測値とモデル予測値から残差を計算
resid_1 <- residuals(fit_1) 
resid_24 <- residuals(fit_24)
resid_168 <- residuals(fit_168)

# SARIMAモデルの係数を多項式形式に変換(AO影響度計算に使用)
pars_1 <- coefs2poly(fit_1) 
pars_24 <- coefs2poly(fit_24)
pars_168 <- coefs2poly(fit_168)

# 各周期と対応する変数をリストにまとめる
periods <- c(1, 24, 168)
resid_list <- list(resid_1, resid_24, resid_168)
pars_list <- list(pars_1, pars_24, pars_168)

# 結果保存用リストの作成
AO_coefhat_list <- list()
AO_tstat_list <- list()


# 周期ごとにAO影響度の算出(論文2.2節)
for (i in seq_along(periods)) {
  resid_ <- resid_list[[i]]
  pars_ <- pars_list[[i]]
  
  resid_imp <- resid_
  ar <- pars_$arcoefs
  ma <- pars_$macoefs
  n <- length(resid_)
  
  picoefs <- c(1, ARMAtoMA(-ma, -ar, n - 1))
  ao.xy <- stats::filter(x = c(resid_, rep(0, n - 1)), filter = rev(picoefs), method = "convolution", sides = 1)
  ao.xy <- as.vector(ao.xy[-seq_along(picoefs[-1])])
  xxinv <- rev(1 / cumsum(picoefs^2))
  coefhat <- ao.xy * xxinv
  
  # σ は MAD ベースの頑健推定（1.483×median(|e - median(e)|)）
  sigma <- 1.483 * quantile(abs(resid_ - quantile(resid_, probs = 0.5, na.rm = TRUE)), probs = 0.5, na.rm = TRUE)
  tstat <- coefhat / (sigma * sqrt(xxinv))
  
  # NA を暫定補完して再計算
  idna <- which(is.na(resid_))
  if (length(idna)) {
    resid_imp[idna] <- mean(resid_, na.rm = TRUE)
  }
  
  ao.xy_i <- stats::filter(x = c(resid_imp, rep(0, n - 1)), filter = rev(picoefs), method = "convolution", sides = 1)
  ao.xy_i <- as.vector(ao.xy_i[-seq_along(picoefs[-1])])
  coefhat_i <- ao.xy_i * xxinv
  sigma_i <- 1.483 * quantile(abs(resid_imp - quantile(resid_imp, probs = 0.5, na.rm = TRUE)), probs = 0.5, na.rm = TRUE)
  AO_tstat <- coefhat_i / (sigma_i * sqrt(xxinv))
  
  # AO影響度(coefhat)とt統計量(tstat)の算出
  AO_coefhat_list[[i]] <- coefhat_i
  AO_tstat_list[[i]] <- AO_tstat
}

# 周期別に結果を格納
AO_coefhat1 <- AO_coefhat_list[[1]]
AO_tstat1   <- AO_tstat_list[[1]]
AO_coefhat24 <- AO_coefhat_list[[2]]
AO_tstat24   <- AO_tstat_list[[2]]
AO_coefhat168 <- AO_coefhat_list[[3]]
AO_tstat168   <- AO_tstat_list[[3]]





#### 4.外れ値検出結果のデータフレーム化 ####
df <- NULL
df <- data.frame(data[,1],date = time(datats1),data$value,datats1,
                    fittedval_1,fittedval_24,fittedval_168,
                    AO_coefhat1,AO_coefhat24,AO_coefhat168,
                    AO_tstat1,AO_tstat24,AO_tstat168) #元データとNA置換後

colnames(df) <- c("date","ind", "val","TF_val", "fit_val1","fit_val24","fit_val168",
                     "AO_coefhat1","AO_coefhat24","AO_coefhat168",
                     "AO_tstat1","AO_tstat24","AO_tstat168")

df$low_th1 <- NA
df$low_th24 <- NA
df$low_th168 <- NA
df$up_th1 <- NA
df$up_th24 <- NA
df$up_th168 <- NA



# パーセンタイルによるAO判定(上限99.5%・下限0.5%のパーセンタイル)
low_per <- 0.005
up_per <- 1-low_per

AO_coefhat <- AO_coefhat1
low_th <- NULL
up_th <- NULL
low_th <- quantile(AO_coefhat, low_per, na.rm = TRUE)
up_th <- quantile(AO_coefhat, up_per, na.rm = TRUE)
df$low_th1 <- low_th
df$up_th1 <- up_th

AO_coefhat <- AO_coefhat24
low_th <- NULL
up_th <- NULL
low_th <- quantile(AO_coefhat, low_per, na.rm = TRUE)
up_th <- quantile(AO_coefhat, up_per, na.rm = TRUE)
df$low_th24 <- low_th
df$up_th24 <- up_th

AO_coefhat <- AO_coefhat168
low_th <- NULL
up_th <- NULL
low_th <- quantile(AO_coefhat, low_per, na.rm = TRUE)
up_th <- quantile(AO_coefhat, up_per, na.rm = TRUE)
df$low_th168 <- low_th
df$up_th168 <- up_th




# coefhatパーセンタイルによる外れ値と欠損値判定(論文2.3節)
df$imp_Judge1 <- ifelse(is.na(df$TF_val), "TF",
                           ifelse(df$AO_coefhat1 < df$low_th1 |
                                    df$AO_coefhat1 > df$up_th1,"AO", NA))

df$imp_Judge24 <- ifelse(is.na(df$TF_val), "TF",
                            ifelse(df$AO_coefhat24 < df$low_th24 |
                                     df$AO_coefhat24 > df$up_th24,"AO", NA))

df$imp_Judge168 <- ifelse(is.na(df$TF_val), "TF",
                             ifelse(df$AO_coefhat168 < df$low_th168 |
                                      df$AO_coefhat168 > df$up_th168,"AO", NA))



# 3つともAO/TFの箇所（論文2.4節）
df <- df %>%
  mutate(imp_Judge_all = case_when(
    imp_Judge1 == "AO" & imp_Judge24 == "AO" & imp_Judge168 == "AO" ~ "AO",
    imp_Judge1 == "TF" | imp_Judge24 == "TF" | imp_Judge168 == "TF" ~ "TF",
    TRUE ~ NA_character_
  ))

# 外れ値(TF,AO)検出数の確認
sum(df$imp_Judge_all == "TF", na.rm = TRUE) #0以下とTFで除外した数
sum(df$imp_Judge_all == "AO", na.rm = TRUE) #3周期AO検出数





#### 5.外れ値の補完処理(線形補間,論文2章の最後) ####
# 補完用のデータ列を追加(TF後の観測値)
df$imp_val <- df$TF_val

# 0.5パーセンタイルでの外れ値検出結果がAOの場合NAに置換
df$imp_val[df$imp_Judge_all == "AO"] <- NA

# 補完用のデータ
imp_data <- NULL
imp_data <- df$imp_val
# imp_datats <- ts(imp_data,frequency = 1) 

# 補完用のデータフレーム作成
imp_ind <- sort(which(is.na(imp_data))) #is.naでNA判定(TRUE/FALSE)→whichでNAの位置取得
imp_df <- NULL # 補完用のデータフレーム作成
imp_df <- data.frame(date = df$date,ind = as.numeric(df$ind), val = df$val, 
                     imp_val = as.numeric(imp_data), m_Judge = is.na(imp_data), line = NA, sline = NA,
                     ARIMA1 = NA , ARIMA2 = NA)
# 欠損値が単発か連続か判定し4列目に入力
imp_df$m_Judge　<- ifelse(imp_df[,5] & !lag(imp_df[,5], default = FALSE) & !lead(imp_df[,5], default = FALSE), 1, 
                         ifelse(imp_df[,5] & (lag(imp_df[,5], default = FALSE) | lead(imp_df[,5], default = FALSE)), 2, NA))

# 外れ値を補完する
line_data <- imp_data
time_index <- 1:length(line_data)
valid_points <- !is.na(line_data)
x <- time_index[valid_points]
y <- line_data[valid_points]
interpolated <- approx(x, y, xout = time_index)
line_data[!valid_points] <- interpolated$y[!valid_points]
imp_df$line <- line_data 

# 補完結果と外れ値検出結果のデータフレームを統合
df <- cbind(df, imp_df[,5:6])





#### 6.データフレームのグラフ化 ####
plot <- plot_ly() %>%
  add_trace(data = df, x = ~date, y = ~val, name = '元', type = 'scatter',mode = 'lines',
            line = list(color = "black")) %>%
  add_trace(data = df, x = ~date, y = ~line, name = '線形補間', type = 'scatter',mode = 'lines',
            line = list(color = "blue")) %>%
  # ARIMAモデルのfitted結果出力用のコード
  # add_trace(data = df, x = ~date, y = ~fit_val1, name = 'fre1', type = 'scatter',mode = 'lines') %>%
  # add_trace(data = df, x = ~date, y = ~fit_val24, name = 'fre24', type = 'scatter',mode = 'lines') %>%
  # add_trace(data = df, x = ~date, y = ~fit_val168, name = 'fre168', type = 'scatter',mode = 'lines') %>%

  # 補完箇所のみをプロット
  add_markers(data = df[df$imp_Judge_all %in% c("AO", "TF"), ] , x = ~date, y = ~line, mode = 'markers',
              name = '補完', marker = list(size = 5, color = 'blue', symbol = "circle")) %>%
  # 外れ値検出箇所をAOとしてプロット
  add_markers(data = df[df$imp_Judge_all == "AO", ] , x = ~date, y = ~val, mode = 'markers', 
              name = 'AO', marker = list(size = 8, color = 'red', symbol = "diamond")) %>%
  add_text(data = df[df$imp_Judge_all == "AO", ], x = ~date, y=~val, text = ~imp_Judge_all, 
           textposition='top center') %>%
  # 前処理で除外した箇所をTFとしてプロット  
  add_markers(data = df[df$imp_Judge_all == "TF", ] , x = ~date, y = ~val, mode = 'markers', 
              name = 'TF', marker = list(size = 8, color = 'orange', symbol = "diamond")) %>%
  add_text(data = df[df$imp_Judge1 == "TF", ], x = ~date, y=~val, text = ~imp_Judge_all, 
           textposition='top center') %>%
  layout(
    xaxis = list(title = "日時",titlefont = list(size = 20),
                 tickfont = list(size = 20),
                 tickformat = "%H:%M<br>%m/%d/%a<br>%Y"), #x軸の文字サイズ,デフォ13
    yaxis = list(title = "電力消費量[kWh]",　titlefont = list(size = 20),
                 tickfont = list(size = 20))   #y軸の文字サイズ,デフォ13
  )
plot




#### 7.ファイル保存 ####
# ファイル名作成
filename_3fre <- paste0("ch",year(start_date), "log-Tuk(k3)_percentile-coef(ARIMA1-24-168)")
# 表示したグラフをhtml形式で保存
saveWidget(plot, paste0(filename_3fre, ".html"))
# データフレームをcsv形式で保存
write.csv(df, paste0(filename_3fre, ".csv"), row.names = FALSE, fileEncoding = "UTF-8")





#### 終了 ####
print(paste0(year(start_date_input), "年度の処理が終了しました"))