import pandas as pd
from pathlib import Path

def generate_sample_data():
    output_dir = Path(__file__).resolve().parent.parent / "sample_data"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "CRM_merge_sample.xlsx"
    
    # Define the 42 original columns or at least the key ones we need
    # To match original 42 columns, we can build a dataframe with all required columns
    columns = [
        "Ngáy cập nhật", "Ngày đăng ký CT", "ActivityId", "Mã CT", "Tên CT", "Địa chỉ", 
        "Loại CT", "Dòng CT", "Quy mô", "Loại hình văn", "Tổng giá trị CT", "Nguồn khai thác", 
        "Doanh thu DK", "Tình trạng ban đầu", "Chủ đầu tư",
        "Tình trạng hiện tại", 
        "Tình hình tiến độ công trình", 
        "Nội dung làm việc, yêu cầu KH & đánh giá", 
        "Kế hoạch lần tới", 
        "Đề xuất"
    ]
    
    # 5 Sample rows showing different classification scenarios
    data = [
        {
            "ActivityId": "ACT_001",
            "Tên CT": "Công trình A",
            "Tình trạng hiện tại": "Đang thi công móng",
            "Tình hình tiến độ công trình": "Công trình khởi động lại chậm, chưa mua hàng.",
            "Nội dung làm việc, yêu cầu KH & đánh giá": "Khách hàng khen đèn âm trần Rạng Đông AT10 9W sáng đẹp, độ bền cao.",
            "Kế hoạch lần tới": "Liên hệ lại tuần sau để báo giá máng đèn.",
            "Đề xuất": "Hỗ trợ catalogue sản phẩm mới."
        },
        {
            "ActivityId": "ACT_002",
            "Tên CT": "Công trình B",
            "Tình trạng hiện tại": "Chuẩn bị lắp thiết bị điện",
            "Tình hình tiến độ công trình": "Đối thủ Điện Quang đang có chiết khấu cao hơn Rạng Đông 5%.",
            "Nội dung làm việc, yêu cầu KH & đánh giá": "Khách hàng phản hồi giá Rạng Đông hơi cao khó cạnh tranh với Điện Quang.",
            "Kế hoạch lần tới": "Giao hàng mẫu sang tuần.",
            "Đề xuất": "Đề xuất cơ chế chiết khấu dự án."
        },
        {
            "ActivityId": "ACT_003",
            "Tên CT": "Công trình C",
            "Tình trạng hiện tại": "Đang hoàn thiện sơn",
            "Tình hình tiến độ công trình": "Công trình hoàn thành 80%, chuẩn bị lấy hàng.",
            "Nội dung làm việc, yêu cầu KH & đánh giá": "Đại lý báo lỗi 2 bóng LED tuýp 1.2m không sáng, cần đổi trả.",
            "Kế hoạch lần tới": "Gặp nhà phân phối thống nhất lịch giao hàng.",
            "Đề xuất": "Gửi sang phòng bảo hành xử lý gấp."
        },
        {
            "ActivityId": "ACT_004",
            "Tên CT": "Công trình D",
            "Tình trạng hiện tại": "Hoạt động bình thường",
            "Tình hình tiến độ công trình": "",
            "Nội dung làm việc, yêu cầu KH & đánh giá": "",
            "Kế hoạch lần tới": "",
            "Đề xuất": ""
        }
    ]
    
    # Create DataFrame filling other columns with None
    df = pd.DataFrame(data, columns=columns)
    df.to_excel(output_path, index=False)
    print(f"[OK] Generated CRM mock sample excel at: {output_path}")

if __name__ == "__main__":
    generate_sample_data()
