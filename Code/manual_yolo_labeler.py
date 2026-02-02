import cv2
import os

IMAGE_DIR = r"C:\Drive_d\Python\F-AI\T7\Input\images\val"
LABEL_DIR = r"C:\Drive_d\Python\F-AI\T7\Input\labels\val"

CLASSES = {
    "l": 0,  # logo
    "h": 1,  # header_block
    "t": 2,  # table
    "g": 3,  # total_block
}

os.makedirs(LABEL_DIR, exist_ok=True)

# Screen parameters (portrait-friendly)
SCREEN_H = 900   # fit to screen height
SCREEN_W = 700   # enough for A4 portrait

drawing = False
ix = iy = -1
current_box = None
saved_boxes = []


def mouse_callback(event, x, y, flags, param):
    global ix, iy, drawing, current_box

    if event == cv2.EVENT_LBUTTONDOWN:
        drawing = True
        ix, iy = x, y
        current_box = None

    elif event == cv2.EVENT_MOUSEMOVE and drawing:
        current_box = (ix, iy, x, y)

    elif event == cv2.EVENT_LBUTTONUP:
        drawing = False
        current_box = (ix, iy, x, y)


def to_yolo(box, img_w, img_h):
    x1, y1, x2, y2 = box
    xc = ((x1 + x2) / 2) / img_w
    yc = ((y1 + y2) / 2) / img_h
    w = abs(x2 - x1) / img_w
    h = abs(y2 - y1) / img_h
    return xc, yc, w, h


for image_name in os.listdir(IMAGE_DIR):
    if not image_name.lower().endswith((".png", ".jpg", ".jpeg")):
        continue

    img_path = os.path.join(IMAGE_DIR, image_name)
    label_path = os.path.join(LABEL_DIR, image_name.replace(".png", ".txt"))

    original = cv2.imread(img_path)
    H, W = original.shape[:2]

    # Scale image to fit screen height (portrait)
    scale = SCREEN_H / H
    disp_w = int(W * scale)
    disp_h = int(H * scale)

    display = cv2.resize(original, (disp_w, disp_h))

    saved_boxes.clear()
    yolo_labels = []

    cv2.namedWindow(image_name, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(image_name, disp_w, disp_h)
    cv2.setMouseCallback(image_name, mouse_callback)

    print(f"\nLabeling: {image_name}")
    print("[l] logo | [h] header | [t] table | [g] total")
    print("[n] next image")

    while True:
        canvas = display.copy()

        # Draw saved boxes
        for (x1, y1, x2, y2, _) in saved_boxes:
            cv2.rectangle(
                canvas,
                (int(x1 * scale), int(y1 * scale)),
                (int(x2 * scale), int(y2 * scale)),
                (0, 255, 0),
                2
            )

        # Draw current box
        if current_box:
            x1, y1, x2, y2 = current_box
            cv2.rectangle(canvas, (x1, y1), (x2, y2), (255, 0, 0), 2)

        cv2.imshow(image_name, canvas)
        key = cv2.waitKey(1) & 0xFF

        if key == ord("n"):
            break

        elif chr(key) in CLASSES and current_box:
            class_id = CLASSES[chr(key)]

            # Map display coords → original image coords
            dx1, dy1, dx2, dy2 = current_box
            x1 = int(dx1 / scale)
            y1 = int(dy1 / scale)
            x2 = int(dx2 / scale)
            y2 = int(dy2 / scale)

            saved_boxes.append((x1, y1, x2, y2, class_id))

            xc, yc, bw, bh = to_yolo([x1, y1, x2, y2], W, H)
            yolo_labels.append(
                f"{class_id} {xc:.6f} {yc:.6f} {bw:.6f} {bh:.6f}"
            )

            current_box = None
            print(f"Saved class {class_id}")

    cv2.destroyAllWindows()

    if yolo_labels:
        with open(label_path, "w") as f:
            f.write("\n".join(yolo_labels))
