import pandas as pd  # type: ignore
import re
from tkinter import Tk, filedialog, Button, Label, messagebox
from pathlib import Path

# ฟังก์ชันสำหรับตรวจสอบรหัสแม็คโคร
def extract_makro_code(value):
    str_value = str(value).strip()
    if str_value.isdigit():
        return str_value
    match = re.search(r'\d+', str_value)
    if match:
        return match.group(0)
    return None

# ฟังก์ชันสำหรับประมวลผลไฟล์
def process_files():
    input_file_paths = filedialog.askopenfilenames(
        title="เลือกไฟล์ Excel ต้นฉบับ (หลายไฟล์)",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not input_file_paths:
        messagebox.showinfo("ไม่มีไฟล์ถูกเลือก", "กรุณาเลือกไฟล์ Excel")
        return

    home = Path.home()
    output_file_path = str(home) + r'\Documents\Makro\Report'
    Path(output_file_path).mkdir(parents=True, exist_ok=True)
    output_file_path += r'\Processed_PO_Result.xlsx'

    # เตรียม DataFrame สำหรับเก็บผลลัพธ์รวม
    all_results = []

    for input_file_path in input_file_paths:
        try:
            # อ่านข้อมูลตั้งแต่แถวที่ 23 และกำหนด dtype คอลัมน์ที่ 10 เป็นข้อความ
            data = pd.read_excel(
                input_file_path,
                skiprows=23,
                header=None,
                engine='openpyxl',
                dtype={10: str}
            )

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {input_file_path}\n\n{e}")

        try:
            shipping_date_index = data.isin(["วันที่ส่งของ"])[0].tolist().index(True)  # ค้นหาวันที่ส่งของในคอลัมน์แรก
            shipping_date = data.iloc[shipping_date_index][2]  # คอลัมน์ที่สาม
        except:
            print("วันที่ส่งของ shipping date not found")
            shipping_date = None

        # ตั้งชื่อคอลัมน์ใหม่โดยอิงจากจำนวนคอลัมน์จริง
        num_columns = data.shape[1]
        default_columns = [f'Unnamed_{i}' for i in range(num_columns)]

        # เพิ่มชื่อคอลัมน์ที่ต้องการ (ถ้ามี)
        if num_columns >= 5:
            default_columns[4] = 'วันที่สั่งสินค้า'
        if num_columns >= 9:
            default_columns[8] = 'รหัสผู้ผลิต'
        if num_columns >= 10:
            default_columns[10] = 'เลขที่ใบสั่งซื้อ'
        if num_columns >= 11:
            default_columns[11] = 'จำนวนสั่งซื้อ'

        # ใช้ชื่อคอลัมน์ที่ปรับปรุงแล้ว
        data.columns = default_columns
        
        # กรองข้อมูล "เลขที่ใบสั่งซื้อ"
        if 'เลขที่ใบสั่งซื้อ' in data.columns:
            data['เลขที่ใบสั่งซื้อ'] = data['เลขที่ใบสั่งซื้อ'].apply(
                lambda x: x if re.match(r'^\d+\.\d+$', str(x)) else None
            )

        # เติมข้อมูลที่ถูก Merge Cell (Forward Fill)
        if 'วันที่สั่งสินค้า' in data.columns:
            data['วันที่สั่งสินค้า'] = data['วันที่สั่งสินค้า'].ffill()
        if 'รหัสผู้ผลิต' in data.columns:
            data['รหัสผู้ผลิต'] = data['รหัสผู้ผลิต'].ffill()
        if 'เลขที่ใบสั่งซื้อ' in data.columns:
            data['เลขที่ใบสั่งซื้อ'] = data['เลขที่ใบสั่งซื้อ'].ffill()
        if 'Unnamed_1' in data.columns:
            data['Unnamed_1'] = data['Unnamed_1'].ffill()

        # สร้าง DataFrame ผลลัพธ์สำหรับไฟล์นี้
        results = []
        current_product = None
        first_column = True
        previous_order_quantity = ''
        for i, row in data.iterrows():
            if pd.notna(row['Unnamed_1']) and isinstance(row['Unnamed_1'], str) and 'STORE' not in row['Unnamed_1']:
                current_product = row['Unnamed_1']
                current_order_quantity = row['จำนวนสั่งซื้อ'] if 'จำนวนสั่งซื้อ' in data.columns else None

            if pd.notna(row['Unnamed_1']) and 'STORE' in str(row['Unnamed_1']):
                store_info = row['Unnamed_1']
                store_match = re.search(r'STORE\s*(\d{2,3})', store_info)
                store2_match = re.search(r'STORE.*?(\d{2,3})$', store_info)
                qty_match = re.search(r'(\d+)$', store_info)

                store = store_match.group(1) if store_match else None
                store2 = store2_match.group(1) if store2_match else None

                makro_code = row['Unnamed_6']
                total_order_amount = row['จำนวนสั่งซื้อ']

                current_order_quantity_copy = current_order_quantity
                if current_order_quantity != row['จำนวนสั่งซื้อ'] and current_order_quantity == previous_order_quantity:
                    current_order_quantity = ''
                previous_order_quantity = current_order_quantity_copy

                if current_product and store and quantity and makro_code==makro_code:
                    # เพิ่มแถว
                    results.append({
                        'วันที่สั่งสินค้า': row['วันที่สั่งสินค้า'] if 'วันที่สั่งสินค้า' in row and first_column else None,
                        'รหัสผู้ผลิต': row['รหัสผู้ผลิต'] if 'รหัสผู้ผลิต' in row and first_column else None,
                        'เลขที่ใบสั่งซื้อ': row['เลขที่ใบสั่งซื้อ'] if 'เลขที่ใบสั่งซื้อ' in row and first_column else None,
                        'รหัสแม็คโคร': makro_code,
                        'ชื่อสินค้า': current_product,
                        'Store': store,
                        'จำนวนสินค้า': quantity,
                        'รวมจำนวนสั่งซื้อ': current_order_quantity,
                        'วันที่ส่งของ': shipping_date if first_column else None,
                    })
                    first_column = False

        # เพิ่มผลลัพธ์ของไฟล์นี้ลงในผลลัพธ์รวม
        all_results.extend(results)
        break

    # รวมผลลัพธ์ทั้งหมดเป็น DataFrame และบันทึกเป็น Excel
    combined_result_df = pd.DataFrame(all_results)
    required_columns = ['รหัสแม็คโคร', 'Store', 'จำนวนสินค้า', 'รวมจำนวนสั่งซื้อ']
    filtered_result_df = combined_result_df.dropna(subset=required_columns)

    try:
        filtered_result_df.to_excel(output_file_path, index=False)
        messagebox.showinfo("สำเร็จ!", f"การประมวลผลเสร็จสิ้น! ไฟล์บันทึกที่: {output_file_path}")
    except Exception as e:
        messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้\n\n{e}")

# สร้าง GUI
def main():
    # เปิดหน้าต่างสำหรับเลือกหลายไฟล์ Excel
    root = Tk()
    root.title("Makro Order Processor")
    root.geometry("400x200")

    Label(root, text="แอปพลิเคชันประมวลผลใบสั่งซื้อ ", font=("Arial", 14)).pack(pady=10)
    Button(root, text="เลือกไฟล์และประมวลผล", command=process_files, font=("Arial", 12)).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
