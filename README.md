# 📈 HOSE Quant Strategy Platform - Web App Kiểm Định Chiến Lược Định Lượng

Ứng dụng web tương tác chuyên sâu được xây dựng trên nền tảng **Streamlit**, dùng để kiểm định và đánh giá tính hiệu quả thực tế của **Chiến lược đầu tư định lượng cổ phiếu HOSE (Phiên bản V4 - Nested Walk-Forward FA/TA + Top N + Dynamic Allocation + Market Regime)** từ dữ liệu lịch sử giá 2021 – 2025.

---

## 🌟 1. Điểm Nổi Bật Của Ứng Dụng

- **Đầy đủ & Khách quan:** Tái hiện chính xác và tối ưu hóa toàn bộ pipeline từ notebook nghiên cứu định lượng:
  - Chia tập dữ liệu thời gian 3 pha: **Core Train**, **Validation** và **Out-of-Sample Final Test**.
  - Tối ưu hóa tĩnh cặp ngưỡng **BUY / EXIT** chung cho thị trường trên Core Train bằng k-fold cross-validation.
  - Tái chọn lọc cổ phiếu định kỳ (**Nested Walk-Forward**) theo quý dựa trên điểm tổng hợp đa nhân tố (TA Score 0-10, Trailing Sharpe, Relative Strength, Volatility, Liquidity, Momentum).
  - Phân bổ tỷ trọng động (**Dynamic Allocation**) theo tháng bằng các trường phái tối ưu danh mục hiện đại: **Equal Weight (EW)**, **Markowitz MPT với Ledoit-Wolf Covariance Shrinkage**, **Minimum Variance (MinVar)**, **Hierarchical Risk Parity (HRP)**, và **Blend 50/50**.
  - Bộ lọc xu hướng thị trường (**Market Regime Overlay - EMA200 + Độ dốc**) giúp phòng vệ, chuyển tài sản về tiền mặt trong các pha suy thoái (Downtrend / Bear Market).
  - Thí nghiệm quản trị tiền mặt linh hoạt (**Cash Policy: Cash 40% / Original / Full Equity**).
- **Trải nghiệm người dùng mượt mà:**
  - Tự động nhận diện dữ liệu Excel sẵn có trong thư mục hoặc cho phép người dùng tải lên file dữ liệu mới (`.xlsx`, `.xls`, `.csv`).
  - Hỗ trợ cả **Giá điều chỉnh (Adjusted Close)** và **Giá khớp lệnh thực tế (Raw Close)**.
  - Biểu đồ tương tác thời gian thực với **Plotly** (Zoom, Pan, Tooltip, Hover date): Đường cong NAV, Sụt giảm sụt giá (Underwater Drawdowns), Tỷ trọng tiền mặt theo thời gian.
  - Bộ soi chỉ báo kỹ thuật chuyên sâu (TA Screener) cho từng mã trong rổ 98 cổ phiếu HOSE.
  - Xuất toàn diện kết quả kiểm định sang file **Excel đa trang tính (.xlsx)** chỉ với 1 cú click.

---

## 📁 2. Cấu Trúc Thư Mục Repository

```text
├── app.py                                                   # Ứng dụng web Streamlit chính
├── requirements.txt                                         # Danh sách thư viện Python cần thiết
├── README.md                                                # Tài liệu hướng dẫn cài đặt & triển khai
└── HOSE - 01.01.2021 - 31.12.2025 (98 ma bao gồm VNINDEX).xlsx # File dữ liệu mẫu lịch sử giá HOSE
```

---

## 🚀 3. Hướng Dẫn Chạy Cục Bộ (Local Machine)

### Bước 1: Chuẩn bị môi trường Python
Khuyến nghị sử dụng **Python 3.9 - 3.12**.

### Bước 2: Cài đặt các thư viện cần thiết
Mở Terminal hoặc Command Prompt / PowerShell tại thư mục dự án và chạy lệnh:

```bash
pip install -r requirements.txt
```

### Bước 3: Khởi chạy ứng dụng Streamlit
Chạy lệnh sau:

```bash
streamlit run app.py
```

Trình duyệt web sẽ tự động mở tại địa chỉ `http://localhost:8501`.

---

## ☁️ 4. Hướng Dẫn Đưa Lên GitHub & Triển Khai Miễn Phí Trên Streamlit Cloud

### Bước 1: Tạo Git Repository và đẩy mã nguồn lên GitHub
1. Mở Terminal tại thư mục này và khởi tạo git:
   ```bash
   git init
   git add app.py requirements.txt README.md "HOSE - 01.01.2021 - 31.12.2025 (98 ma bao gồm VNINDEX).xlsx"
   git commit -m "Initial commit: Deploy HOSE Quant Strategy Streamlit App"
   ```
2. Tạo một repository mới trên GitHub (ví dụ: `hose-quant-strategy`).
3. Liên kết và đẩy code lên:
   ```bash
   git remote add origin https://github.com/<tai-khoan-github-cua-ban>/hose-quant-strategy.git
   git branch -M main
   git push -u origin main
   ```

### Bước 2: Triển khai trên Streamlit Community Cloud
1. Truy cập vào **[share.streamlit.io](https://share.streamlit.io)** và đăng nhập bằng tài khoản GitHub của bạn.
2. Nhấn nút **"New app"** (hoặc **"Create app"**).
3. Điền thông tin cấu hình:
   - **Repository:** `<tai-khoan-github-cua-ban>/hose-quant-strategy`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Nhấn **"Deploy!"**.
5. Sau 1 - 2 phút cài đặt các gói trong `requirements.txt`, web app của bạn sẽ chính thức hoạt động trực tuyến với đường link công khai dạng `https://<ten-app>.streamlit.app`!

---

## 📊 5. Cấu Trúc Các Tab Chức Năng Trong Ứng Dụng

| Tab | Tên Chức Năng | Nội Dung & Giá Trị Phân Tích |
| :--- | :--- | :--- |
| **Tab 1** | **📈 Final Test (OOS)** | Hiển thị kết quả kiểm định Out-of-Sample độc lập của mô hình Champion so với VN-Index và Static Buy & Hold; biểu đồ NAV chuẩn hóa; biểu đồ mức sụt giảm (Max Drawdown); bảng so sánh đầy đủ các chỉ số: CAGR, Sharpe, Sortino, Calmar, Volatility, Turnover. |
| **Tab 2** | **🏆 Validation & Champion** | Đấu loại các họ mô hình phân bổ tỷ trọng (EW, MPT-LW, MinVar-LW, HRP, Blend50) trên tập Validation; công thức tính điểm Validation Score và công bố mô hình Champion chiến thắng. |
| **Tab 3** | **🔄 Top N & Phân Bổ** | Lịch tái lựa chọn danh mục Top N theo từng quý; bảng phân bổ chi tiết các mã cổ phiếu và biểu đồ tần suất xuất hiện của các mã dẫn đầu. |
| **Tab 4** | **🛡️ Market Regime & Tiền Mặt** | Nhận diện chu kỳ thị trường (Risk-On / Risk-Off) qua đường trung bình EMA 200 và độ dốc xác nhận; mô phỏng hiệu quả của chính sách tiền mặt `cash40` (giữ tối đa 40% tiền mặt) so với `original` và `full`. |
| **Tab 5** | **📑 Nhật Ký Lệnh & Audit** | Danh sách khớp lệnh chi tiết từng ngày (Date, Ticker, BUY/SELL, Khối lượng, Giá khớp, Phí); bảng **Leakage Audit** cam kết không sử dụng dữ liệu tương lai (No Lookahead Bias); tính năng tải xuống báo cáo Excel (.xlsx). |
| **Tab 6** | **🔍 Khám Phá Dữ Liệu & TA** | Bảng kiểm tra chất lượng dữ liệu của 98 mã cổ phiếu; công cụ trực quan hóa biến động giá, khối lượng và biểu đồ điểm kỹ thuật TA Score (0-10) chi tiết cho từng mã được chọn. |

---

## 🔬 6. Phương Pháp Định Lượng & Quy Tắc Phòng Ngừa Rò Rỉ Dữ Liệu (Leakage Audit)

1. **Điểm Kỹ Thuật (TA Score 0-10):**
   - **Trend (0 - 5 điểm):** $Close > EMA_{50}$ (2đ), $EMA_{50} > EMA_{200}$ (2đ), $MACD > MACD_{Signal}$ (1đ).
   - **Momentum (0 - 3 điểm):** $RSI \in [45, 75]$ hoặc $ROC_{10} > 0$ (2đ), $Close \ge 0.85 \times High_{52W}$ (1đ).
   - **Volume (0 - 1 điểm):** $Volume \ge Volume_{MA20}$ hoặc $OBV > OBV_{EMA20}$ (1đ).
   - **Volatility (0 - 1 điểm):** $ATR_{\%} \le Q90(ATR_{\%})$ (1đ).
2. **Khóa Ngưỡng BUY / EXIT Chung:**
   - Ngưỡng BUY (5, 6, 7) và EXIT (2, 3, 4) được tìm kiếm thông qua Cross-Validation trên tập **Core Train**, tuyệt đối không điều chỉnh trên tập Validation hay Final Test.
3. **Thực Thi Lệnh Thực Tế (Execution Lag):**
   - Tín hiệu tạo ra tại giá đóng cửa $Close(t)$ chỉ được phép phát lệnh và khớp vào giá mở cửa phiên kế tiếp $Open(t+1)$, tính kèm chi phí giao dịch (0.15%) và trượt giá (0.05%).
4. **Covariance Shrinkage (Ledoit-Wolf):**
   - Tránh hiện tượng nhiễu ma trận hiệp phương sai mẫu khi số lượng tài sản và chu kỳ quan sát hạn chế, đảm bảo trọng số MPT và MinVar luôn ổn định và hội tụ.

---

## 📄 7. Giấy Phép & Tác Quyền
Ứng dụng được xây dựng phục vụ mục đích nghiên cứu, học thuật và kiểm định chiến lược đầu tư định lượng trên thị trường chứng khoán Việt Nam.
