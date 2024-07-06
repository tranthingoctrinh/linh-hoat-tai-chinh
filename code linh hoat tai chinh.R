
# Tải các thư viện cần thiết
install.packages(c("ggplot2", "cowplot","dplyr", "scales", "gridExtra"))
library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(openxlsx)
library(plm)
library(scales)
library(ggplot2)
library(gridExtra)
library(cowplot)
# Đặt đường dẫn tới file Excel
file_path <- "D:/k214142094/K214142094.xlsx"
# Đọc dữ liệu từ các sheet trong file Excel
focf <- read_excel(file_path, sheet = "focf", skip = 1, col_names = c("Company", "Date", "FOCF"))
total_equity <- read_excel(file_path, sheet = "total equity", skip = 1, col_names = c("Company", "Date", "Total_Equity"))
cash_equivalents <- read_excel(file_path, sheet = "Cash and Equivalents", skip = 1, col_names = c("Company", "Date", "Cash_Equivalents"))
current_liabilities <- read_excel(file_path, sheet = "Current Liabilities", skip = 1, col_names = c("Company", "Date", "Current_Liabilities"))
total_debt <- read_excel(file_path, sheet = "total debt", skip = 1, col_names = c("Company", "Date", "Total_Debt"))
Net_income <- read_excel(file_path, sheet = "Net Income", skip = 1, col_names = c("Company", "Date", "Net_income"))
# Chuyển đổi cột Date sang kiểu ký tự
focf <- focf %>% mutate(Date = as.character(Date))
total_equity <- total_equity %>% mutate(Date = as.character(Date))
cash_equivalents <- cash_equivalents %>% mutate(Date = as.character(Date))
current_liabilities$Date <- as.character(current_liabilities$Date)
total_debt$Date <- as.character(total_debt$Date)
Net_income$Date <- as.character(Net_income$Date)
# Kết hợp dữ liệu từ các sheet
combined_data <- focf %>%
  left_join(total_equity, by = c("Company", "Date")) %>%
  left_join(cash_equivalents, by = c("Company", "Date")) %>%
  left_join(current_liabilities, by = c("Company", "Date")) %>%
  left_join(total_debt, by = c("Company", "Date"))%>%
  left_join(Net_income, by = c("Company", "Date"))



# Kiểm tra kiểu dữ liệu của các cột
sapply(combined_data, class)

numeric_cols <- c("FOCF", "Total_Equity",  "Cash_Equivalents",  "Current_Liabilities", "Total_Debt","Net_income")
combined_data[numeric_cols] <- lapply(combined_data[numeric_cols], function(x) suppressWarnings(as.numeric(x)))

# Thay thế NA bằng 0 chỉ trong các cột số
filtered_data <- combined_data %>%
  mutate(across(where(is.numeric), ~ifelse(is.na(.), 0, .)))

# Xóa các hàng có giá trị NA
filtered_data <- filtered_data %>%
  na.omit()

# Đếm số lần xuất hiện của mỗi giá trị trong cột "Company"
company_counts <- filtered_data %>%
  count(Company, name = "count")

# Lọc các công ty xuất hiện đúng 9 lần
companies_with_9_occurrences <- company_counts %>%
  filter(count == 9)

# Lọc các hàng chỉ giữ lại các công ty xuất hiện đúng 9 lần
filtered_data <- filtered_data %>%
  filter(Company %in% companies_with_9_occurrences$Company)

# Xem kết quả
head(filtered_data)

# Số lượng quan sát
n_observations <- nrow(filtered_data)
print(paste("Số lượng quan sát:", n_observations))

# Số lượng biến
n_variables <- ncol(filtered_data)
print(paste("Số lượng biến:", n_variables))

# Loại dữ liệu của từng biến
variable_types <- sapply(filtered_data, class)
print("Loại dữ liệu của từng biến:")
print(variable_types)
# Thống kê mô tả định lượng
summary_statistics <- summary(filtered_data)
print(summary_statistics)

# Chuyển đổi cột Date sang kiểu Date và trích xuất năm
filtered_data <- filtered_data %>%
  mutate(Year = year(as.Date(Date)))


# Tính tổng FOCF cho mỗi năm
annual_focf <- filtered_data %>%
  group_by(Year) %>%
  summarise(Total_FOCF = sum(FOCF, na.rm = TRUE))

# Tính tổng Total_Equity cho mỗi năm
annual_total_equity <- filtered_data %>%
  group_by(Year) %>%
  summarise(Total_Equity = sum(Total_Equity, na.rm = TRUE))

# Tính tổng Cash_Equivalents cho mỗi năm
annual_cash_equivalents <- filtered_data %>%
  group_by(Year) %>%
  summarise(Total_Cash_Equivalents = sum(Cash_Equivalents, na.rm = TRUE))

# Tính tổng Current_Liabilities cho mỗi năm
annual_current_liabilities <- filtered_data %>%
  group_by(Year) %>%
  summarise(Total_Current_Liabilities = sum(Current_Liabilities, na.rm = TRUE))


# Vẽ biểu đồ cột cho tổng giá trị FOCF theo năm mà không hiển thị giá trị trên cột
p1 <- ggplot(annual_focf, aes(x = Year, y = Total_FOCF)) +
  geom_bar(stat = "identity", fill = "blue") +
  labs(title = "Tổng giá trị FOCF theo năm", x = "Year", y = "Total FOCF") +
  scale_x_continuous(breaks = seq(min(annual_focf$Year), max(annual_focf$Year), by = 1)) +
  scale_y_continuous(labels = scales::comma) +
  theme_minimal()

print(p1)


# Vẽ biểu đồ cột cho tổng giá trị Total Equity theo năm mà không hiển thị giá trị trên cột
p2 <- ggplot(annual_total_equity, aes(x = Year, y = Total_Equity)) +
  geom_bar(stat = "identity", fill = "red") +
  labs(title = "Tổng giá trị Total Equity theo năm", x = "Year", y = "Total Equity") +
  scale_x_continuous(breaks = seq(min(annual_total_equity$Year), max(annual_total_equity$Year), by = 1)) +
  scale_y_continuous(labels = scales::comma) +
  theme_minimal()

print(p2)

# Vẽ biểu đồ cột cho tổng giá trị Cash Equivalents theo năm mà không hiển thị giá trị trên cột
p3 <- ggplot(annual_cash_equivalents, aes(x = Year, y = Total_Cash_Equivalents)) +
  geom_bar(stat = "identity", fill = "green") +
  labs(title = "Tổng giá trị Cash Equivalents theo năm", x = "Year", y = "Total Cash Equivalents") +
  scale_x_continuous(breaks = seq(min(annual_cash_equivalents$Year), max(annual_cash_equivalents$Year), by = 1)) +
  scale_y_continuous(labels = scales::comma) +
  theme_minimal()

print(p3)


# Vẽ biểu đồ cột cho tổng giá trị Current Liabilities theo năm mà không hiển thị giá trị trên cột
p4 <- ggplot(annual_current_liabilities, aes(x = Year, y = Total_Current_Liabilities)) +
  geom_bar(stat = "identity", fill = "purple") +
  labs(title = "Tổng giá trị Current Liabilities theo năm", x = "Year", y = "Total Current Liabilities") +
  scale_x_continuous(breaks = seq(min(annual_current_liabilities$Year), max(annual_current_liabilities$Year), by = 1)) +
  scale_y_continuous(labels = scales::comma) +
  theme_minimal()

print(p4)



# Chuyển đổi dữ liệu thành dạng dài
melted_data <- filtered_data %>%
  pivot_longer(cols = c(FOCF, Total_Equity,  Cash_Equivalents, Current_Liabilities,Total_Debt,Net_income),
               names_to = "Variable",
               values_to = "Value")

# Biểu đồ boxplot của các biến
box_plot <- ggplot(melted_data, aes(x = Variable, y = Value)) +
  geom_boxplot(fill = "skyblue", color = "blue") +
  labs(title = "Biểu đồ Boxplot của các biến", x = "Biến", y = "Giá trị") +
  theme_minimal()
print(box_plot)

# Biểu đồ Heatmap
correlation_matrix <- cor(filtered_data[, c("FOCF", "Total_Equity", "Cash_Equivalents",  "Current_Liabilities","Total_Debt","Net_income")])

# Chuyển đổi ma trận tương quan thành dạng dữ liệu tidy
correlation_data <- as.data.frame(as.table(correlation_matrix))
names(correlation_data) <- c("Variable1", "Variable2", "Correlation")

heatmap <- ggplot(data = correlation_data, aes(x = Variable1, y = Variable2, fill = Correlation)) +
  geom_tile() +
  scale_fill_gradient2(low = "blue", high = "red", mid = "white", midpoint = 0, limit = c(-1,1), space = "Lab", name="Correlation") +
  labs(title = "Biểu đồ Heatmap của Tương quan giữa các Biến",
       x = "Biến",
       y = "Biến",
       fill = "Hệ số Tương quan") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, size = 10, hjust = 1))

print(heatmap)


# Vẽ biểu đồ hộp cho FOCF theo năm
box_plot <- ggplot(filtered_data, aes(x = factor(Year), y = FOCF)) +
  geom_boxplot(fill = "skyblue", color = "blue") +
  labs(title = "Biểu đồ hộp cho FOCF theo năm", x = "Year", y = "FOCF") +
  theme_minimal()

print(box_plot)


# Tìm outlier trong từng cột sử dụng phương pháp Z-score
find_outliers_zscore <- function(column) {
  mean_val <- mean(column)
  sd_val <- sd(column)
  z_score <- abs((column - mean_val) / sd_val)
  threshold <- 3  # Chọn ngưỡng
  
  outlier_indices <- which(z_score > threshold)
  return(outlier_indices)
}
# Loại bỏ các cột không phải là số hoặc có giá trị trống
filtered_numeric_data <- filtered_data[, sapply(filtered_data, is.numeric)]
filtered_numeric_data <- na.omit(filtered_numeric_data)

# Áp dụng hàm find_outliers_zscore cho từng cột trong DataFrame đã lọc
outlier_indices_list <- lapply(filtered_numeric_data, find_outliers_zscore)

# In ra các chỉ số của outlier trong từng cột
names(outlier_indices_list) <- colnames(filtered_numeric_data)
print(outlier_indices_list)


sapply(filtered_data, class)


# Tạo biến Financial Flexibility
filtered_data <- filtered_data %>%
  mutate(Financial_Flexibility = FOCF + Net_income)


# Chuyển đổi dữ liệu thành dạng panel
pdata <- pdata.frame(filtered_data, index = c("Company", "Date"))


# Xóa các hàng có giá trị Infinity hoặc NaN trong pdata
pdata <- pdata[is.finite(pdata$Financial_Flexibility), ]

# Tạo mô hình
fixed_effects_model <- plm(Financial_Flexibility ~ FOCF + Total_Equity + Cash_Equivalents + Current_Liabilities, data = pdata, model = "within")

# Xem kết quả hồi quy
summary(fixed_effects_model)

