import cv2
import mediapipe as mp
import time
import openpyxl
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# class poseDetector():
#     def __init__(self,mode=False,upBody=False,smooth=True,detectionCon=0.5,trackCon=0.5):
#         self.mode=mode
#         self.upBody=upBody
#         self.smooth=smooth
#         self.detectionCon=detectionCon
#         self.trackCon=trackCon

# Set up Google Sheets API credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\\Users\\Shreya Gupta\\Downloads\\eng-digit-408508-a45019e6108a.json', scope)
client = gspread.authorize(creds)

# Open the Google Sheet by title
#spreadsheet = client.open("Data.xlsx")
#worksheet = spreadsheet.Sheet1  # You may need to change the sheet name

url='1qCCXEVrqhrYyPJG-0iGc9RzZRbzIjcUVPxODwqj7iLo'
spreadsheet=client.open_by_key(url)
worksheet = spreadsheet.sheet1 
mpPose = mp.solutions.pose
pose = mpPose.Pose()
mpDraw=mp.solutions.drawing_utils
pTime = 0
cap = cv2.VideoCapture('sample1.mp4')
i=1
landmark=[]
worksheet = spreadsheet.get_worksheet(0) 
while True:
    success, img = cap.read()
    imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = pose.process(imgRGB)
    
    
    if results.pose_landmarks:
        mpDraw.draw_landmarks(img,results.pose_landmarks,mpPose.POSE_CONNECTIONS)
        for id,lm in enumerate(results.pose_landmarks.landmark):
            
            h,w,c=img.shape
            print(id,lm)
            cx,cy=int(lm.x*w),int(lm.y*h)
            cv2.circle(img,(cx,cy),5,(255,0,0),cv2.FILLED)
            #cv2.putText(img, str(id), (cx - 10, cy - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)


            #landmarks = results.pose_landmarks.landmark
            #print(landmarks)
        #Extract landmark values
            landmark_values = [id,lm.x*w ,
                    lm.y*h ,
                    lm.z*c ]
            landmark.append(landmark_values)
               
            
        
            
        #Append data to Google Sheet
            
             # Replace sheet_index with the index of your workshee
            
            worksheet.append_row(landmark_values)
            
            

    
    # worksheet = spreadsheet.get_worksheet(0) 
    # for row_values in landmark:
    #     worksheet.append_row(row_values)
    time.sleep(20)
    cv2.imshow("Image", img)  
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    
    cv2.putText(img, str(int(fps)), (70, 50), cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)
    
    cv2.waitKey(50)

# Close the workbook (not necessary in this case, as the script is running indefinitely)
# wBook.save('Data.xlsx')
