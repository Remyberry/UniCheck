import ast
from pyzbar.pyzbar import decode 
import cv2


cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
ret, frame = cap.read()
while True:
    ret, frame = cap.read()
    if ret:
        frame = cv2.flip(frame, 1)
        cv2.imshow("Webcam", frame)
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        # Detect QR codes in the frame
        qrcodes = decode(gray)

        if qrcodes:
            # Extract data from decoded object
            data = qrcodes[0].data.decode("utf-8")
            qr_data = data  # Detect objects in the frame

            if qr_data is not None:
                qr_data_str = ast.literal_eval(qr_data)
                qr_data_dict = ast.literal_eval(qr_data_str)
                for key, value in qr_data_dict.items():
                    print(f"{key}: {value}")
                output_string = ", ".join(f"{key}: {value}" for key, value in qr_data_dict.items())
                print(output_string)


    if cv2.waitKey(1) & 0xFF == ord('q'):
        break


# qr_data = '{"StudentID": "21-0015", "StudentName": "Jerry Santos", "Course": "BSCS-403", "Gmail": "jmsantos@cca.edu.ph"}'

# qr = ast.literal_eval(qr_data)
# print(type(qr))