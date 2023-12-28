import cv2
import mediapipe as mp
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Set up Google Sheets API credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\\Users\\Shreya Gupta\\Downloads\\eng-digit-408508-a45019e6108a.json', scope)
client = gspread.authorize(creds)

# Open the Google Sheet by title
url='1qCCXEVrqhrYyPJG-0iGc9RzZRbzIjcUVPxODwqj7iLo'
spreadsheet=client.open_by_key(url)
worksheet = spreadsheet.sheet1 
mpPose = mp.solutions.pose
pose = mpPose.Pose()
mpDraw = mp.solutions.drawing_utils
pTime = 0
cap = cv2.VideoCapture('sample1.mp4')
i = 1
landmark = []

# Create the worksheet outside the loop
worksheet = spreadsheet.get_worksheet(0)

while True:
    success, img = cap.read()
    imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = pose.process(imgRGB)
    
    if results.pose_landmarks:
        mpDraw.draw_landmarks(img, results.pose_landmarks, mpPose.POSE_CONNECTIONS)
        for id, lm in enumerate(results.pose_landmarks.landmark):
            h, w, c = img.shape
            cx, cy = int(lm.x * w), int(lm.y * h)
            cv2.circle(img, (cx, cy), 5, (255, 0, 0), cv2.FILLED)
            #cv2.putText(img, str(id), (cx - 10, cy - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)

            # Extract landmark values
            landmark_values = [id, lm.x * w, lm.y * h, lm.z * c]
            landmark.append(landmark_values)

    # Append data to Google Sheet
    for row_values in landmark:
        worksheet.append_row(row_values)

    cv2.imshow("Image", img)  
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    
    cv2.putText(img, str(int(fps)), (70, 50), cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)
    
    cv2.waitKey(1)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release the video capture object and close all OpenCV windows
cap.release()
cv2.destroyAllWindows()
