import cv2
import argparse
import torch
import csv
from pyzbar.pyzbar import decode    #QR CODE SCANNER
from ultralytics import YOLO
import supervision as sv

device = 'cuda' if torch.cuda.is_available() else 'cpu'
print(f'Using device: {device}')

def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="YOLOv8 live")
    parser.add_argument(
        "--webcam-resolution", 
        default=[1280, 720], 
        nargs=2, 
        type=int
    )
    args = parser.parse_args()
    return args

def appendFiles(data, csv_file_path):
    with open(csv_file_path, 'a', newline='', encoding="utf-8") as csvfile:                 #Writing a csv file with UTF-8 encoding
        writer = csv.writer(csvfile)                                                        #New Instance of csv.writer
        writer.writerow(["StudentID", "StudentName", "Email", "Section"])                   #Assigning columns
        for item in data:
            writer.writerow([item['StudentID'],item['StudentName'],item['Email'],item['Section'],item['Comment']])

def main():
    args = parse_arguments()
    studentQRData = []
    frame_width, frame_height = args.webcam_resolution

    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, frame_width)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, frame_height)

    model = YOLO('custom_model.pt')

    box_annotator = sv.BoxCornerAnnotator(thickness=2)


    while True:
        ret, frame = cap.read()
        if ret:

            result = model(frame, conf=0.50, show_labels=True, show_conf=True)[0]

            detections = sv.Detections.from_ultralytics(result)
            labels = [f"{class_name} {confidence:.2f}"
                for class_name, confidence in zip(detections['class_name'], detections.confidence)]
            label_annotator = sv.LabelAnnotator(text_position=sv.Position.TOP_CENTER,
                                        text_scale=0.5,
                                        text_thickness=1,
                                        text_padding=10)

            frame = box_annotator.annotate(scene=frame, detections=detections)
            frame = label_annotator.annotate(scene=frame,detections=detections,labels=labels)
            
            # Convert frame to grayscale for QR code detection
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            # Decode QR code
            qr_data = decode(gray)

            # Check if QR code is detected
            if qr_data:
                # Extract data from decoded object
                data = qr_data[0].data.decode("utf-8")
                print("Decoded Data:", data)

            
            cv2.imshow("Webcam", frame)

            if cv2.waitKey(1) == ord('q'):
                break

    cap.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()