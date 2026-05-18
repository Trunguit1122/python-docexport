# python-docexport

Tạo mẫu Word văn bản hành chính theo Nghị định 30/2020/NĐ-CP bằng `python-docx`.

## Cài đặt

```bash
python3 -m pip install python-docx
```

## Cách dùng

Liệt kê loại văn bản hỗ trợ:

```bash
python3 main.py --list
```

Tạo đủ 25 mẫu vào thư mục `output_nd30`:

```bash
python3 main.py --all
```

Tạo một loại mẫu:

```bash
python3 main.py --type cong_van
python3 main.py --type quyet_dinh --out-dir output_nd30
```

## Ghi chú

- Các file mẫu sinh sẵn nằm trong `output_nd30/`.
- Nội dung trong dấu ngoặc vuông là placeholder để thay thế khi soạn văn bản thật.
- Mẫu này tập trung vào thể thức/kỹ thuật trình bày. Khi phát hành chính thức vẫn cần kiểm tra thẩm quyền, số/ký hiệu, nơi nhận, chữ ký, con dấu/chữ ký số và quy chế nội bộ.
