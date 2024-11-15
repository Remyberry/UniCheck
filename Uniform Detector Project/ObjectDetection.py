import time
import cv2
import torch
import csv
from pyzbar.pyzbar import decode    #QR CODE SCANNER
from ultralytics import YOLO
import supervision as sv

# device = 'cuda' if torch.cuda.is_available() else 'cpu'
# print(f'Using device: {device}')

def appendFiles(data, csv_file_path):
    with open(csv_file_path, 'a', newline='', encoding="utf-8") as csvfile:                 #Writing a csv file with UTF-8 encoding
        writer = csv.writer(csvfile)                                                        #New Instance of csv.writer
        writer.writerow(["StudentID", "StudentName", "Section", "Email", "Remarks", "Capture"])                   #Assigning columns
        for item in data:
            writer.writerow([item['StudentID'], item['StudentName'], item['Section'], item['Email'], item['Remarks'], item['Capture']])

def main():
    cap = cv2.VideoCapture(1, cv2.CAP_DSHOW)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
    cap.set(cv2.CAP_PROP_BUFFERSIZE, 2)
    model = YOLO("custom_model.pt")

    class_colors = {
        "Not_In_Uniform": sv.Color(255, 0, 0),  # Red
        "PE_NSTP": sv.Color(0, 255, 0),  # Green
        "Person": sv.Color(255, 165, 0),  # Orange
        "In_Uniform": sv.Color(0, 255, 0)  # Green
    }
    detection_counter = 0

    while True:
        ret, frame = cap.read()
        if not ret:
            break
        frame = cv2.flip(frame, 1)
        #Detection block
        if detection_counter % 3 == 0:
            result = model(frame, conf=0.50, show_labels=True, show_conf=True)[0]
            detections = sv.Detections.from_ultralytics(result)
            label_annotator = sv.LabelAnnotator(text_position=sv.Position.TOP_CENTER,
                                        text_scale=0.5,
                                        text_thickness=1,
                                        text_padding=10)
            box_annotator = sv.BoxCornerAnnotator(thickness=2)


            new_labels = []
            for class_name, confidence in zip(detections['class_name'], detections.confidence):
                if class_name == "notInUniform":
                    new_labels.append("Not_In_Uniform")
                elif class_name == "pe-ntstp":
                    new_labels.append("PE_NSTP")
                elif class_name == "person":
                    new_labels.append("Person")
                elif class_name == "uniforms":
                    new_labels.append("In_Uniform")
                # color = class_colors.get(new_labels[-1], (0, 0, 0))  # Default to black
                new_labels[-1] = f"{new_labels[-1]} ({confidence:.2f})"
                # label_annotator.color = color
                # box_annotator.color = color
            # labels = [f"{class_name} {confidence:.2f}"
            #     for class_name, confidence in zip(detections['class_name'], detections.confidence)]
            
        detection_counter += 1
        frame = box_annotator.annotate(scene=frame, detections=detections)
        frame = label_annotator.annotate(scene=frame,detections=detections,labels=new_labels)

        cv2.imshow("Webcam", frame)

         # Convert frame to grayscale for QR code detection
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # Decode QR code
        qr_data = decode(gray)

        # Check if QR code is detected
        if qr_data:
            # Extract data from decoded object
            data = qr_data[0].data.decode("utf-8")
            print("Decoded Data:", data)
        if cv2.waitKey(1) == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()