from ultralytics import YOLO

model = YOLO("yolov8s.pt")  # nano = best for small dataset

model.train(
    data=r"C:/Drive_d/Python/F-AI/T7/Code/data.yaml",
    imgsz=1024,
    epochs=80,
    batch=2,
    project="invoice_layout",
    name="yolov8s_layout_v1"
)
