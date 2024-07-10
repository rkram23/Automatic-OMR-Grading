import cv2
import numpy as np
import utlis
import openpyxl
import keyboard
##################################################
path = "1.jpg"
widthImg=550
heightImg=550
questions = 5
choices = 5
ans = [1,2,0,2,4]
webcamfeed = False
cameraNo = 0
##################################################

cap = cv2.VideoCapture(cameraNo)
cap.set(10,150)

# Initialize Excel workbook and worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = "Grades"
worksheet.append(["Name", "Email", "Score"])


# Function to save data to Excel
def save_to_excel(name, email, score):
    worksheet.append([name, email, score])
    workbook.save("grades.xlsx")


# Function to send email
def send_email(name, email, score):
    # Configure your email settings
    smtp_server = ""  # Replace with your SMTP server address
    smtp_port = 587  # Replace with your SMTP server port
    sender_email = ""  # Replace with your email address
    sender_password = getpass.getpass("Enter your email password: ")  # Prompt for your email password

    receiver_email = email
    subject = "Your Test Score"
    message = f"Hello {name},\n\nYour test score is: {score}%"

    # Create and send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Subject"] = subject

        msg.attach(MIMEText(message, "plain"))

        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Function to schedule email for later
def schedule_email(date_time, name, email, score):
    current_time = datetime.datetime.now()
    time_difference = (date_time - current_time).total_seconds()

    if time_difference <= 0:
        print("Invalid date and time. The specified time has already passed.")
        return

    def send_email_later():
        nonlocal name, email, score
        send_email(name, email, score)

    # Schedule the email to be sent in the future
    timer = threading.Timer(time_difference, send_email_later)
    timer.start()


count = 0
name = ""
email = ""
while True:
    if webcamfeed:
        success, img = cap.read()
    else :
        img=cv2.imread(path)
    # Preprocessing
    img = cv2.resize(img, (widthImg, heightImg))
    imgContours = img.copy()
    imgFinal = img.copy()
    imgBiggestContours = img.copy()
    imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    imgBlur = cv2.GaussianBlur(imgGray, (5, 5), 1)
    imgCanny = cv2.Canny(imgBlur, 10, 50)


    try:
        # Finding All Contours
        contours, hierarchy = cv2.findContours(imgCanny, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
        cv2.drawContours(imgContours, contours, -1, (0, 255, 0), 10)

        # Finding Rectangles
        rectCon = utlis.rectContour(contours)
        biggestContour = utlis.getCornerPoints(rectCon[0])

        gradePoints = utlis.getCornerPoints(rectCon[1])
        # print(len(biggestContour))
        if biggestContour.size != 0 and gradePoints.size != 0:
            cv2.drawContours(imgBiggestContours, biggestContour, -1, (0, 255, 0), 20)
            cv2.drawContours(imgBiggestContours, gradePoints, -1, (255, 0, 0), 20)

            biggestContour = utlis.reorder(biggestContour)
            gradePoints = utlis.reorder(gradePoints)

            pt1 = np.float32(biggestContour)
            pt2 = np.float32([[0, 0], [widthImg, 0], [0, heightImg], [widthImg, heightImg]])
            matrix = cv2.getPerspectiveTransform(pt1, pt2)
            imgWarpColored = cv2.warpPerspective(img, matrix, (widthImg, heightImg))

            ptG1 = np.float32(gradePoints)
            ptG2 = np.float32([[0, 0], [325, 0], [0, 150], [325, 150]])
            matrixG = cv2.getPerspectiveTransform(ptG1, ptG2)
            imgGradeDisplay = cv2.warpPerspective(img, matrixG, (325, 150))
            # cv2.imshow("Grade",imgGradeDisplay)

            # Apply Threshold
            imgWarpGray = cv2.cvtColor(imgWarpColored, cv2.COLOR_BGR2GRAY)
            imgThresh = cv2.threshold(imgWarpGray, 170, 255, cv2.THRESH_BINARY_INV)[1]

            boxes = utlis.splitBoxes(imgThresh)
            # cv2.imshow("Test",boxes[2])
            # print(cv2.countNonZero(boxes[1]),cv2.countNonZero(boxes[2]))

            # Getting Non Zeros pixel values of each box
            myPixelVal = np.zeros((questions, choices))
            countC = 0
            countR = 0

            for image in boxes:
                totalPixels = cv2.countNonZero(image)
                myPixelVal[countR][countC] = totalPixels
                countC += 1
                if (countC == choices):
                    countR += 1
                    countC = 0
            print(myPixelVal)

            # Finding Index Values of the Markings
            myIndex = []
            for x in range(0, questions):
                arr = myPixelVal[x]
                # print("arr",arr)
                myIndexVal = np.where(arr == np.amax(arr))
                # print(myIndexVal[0])
                myIndex.append(myIndexVal[0][0])
            print(myIndex)

            # Grading
            grading = []
            for x in range(0, questions):
                if ans[x] == myIndex[x]:
                    grading.append(1)
                else:
                    grading.append(0)
            # print(grading)
            score = (sum(grading) / questions) * 100  # Final Grade
            print(score)

            # Display Answers
            imgResult = imgWarpColored.copy()
            imgResult = utlis.showAnswers(imgResult, myIndex, grading, ans, questions, choices)
            imgRawDrawing = np.zeros_like(imgWarpColored)
            imgRawDrawing = utlis.showAnswers(imgRawDrawing, myIndex, grading, ans, questions, choices)
            invMatrix = cv2.getPerspectiveTransform(pt2, pt1)
            imgInvWarp = cv2.warpPerspective(imgRawDrawing, invMatrix, (widthImg, heightImg))

            imgRawGrade = np.zeros_like(imgGradeDisplay)
            cv2.putText(imgRawGrade, str(int(score)) + "%", (60, 100), cv2.FONT_HERSHEY_COMPLEX, 3, (0, 255, 255), 3)
            invMatrixG = cv2.getPerspectiveTransform(ptG2, ptG1)
            imgInvGradeDisplay = cv2.warpPerspective(imgRawGrade, invMatrixG, (widthImg, heightImg))

            imgFinal = cv2.addWeighted(imgFinal, 1, imgInvWarp, 1, 0)
            imgFinal = cv2.addWeighted(imgFinal, 1, imgInvGradeDisplay, 1, 0)

        imgBlank = np.zeros_like(img)
        imageArray = ([img, imgGray, imgBlur, imgCanny],
                      [imgContours, imgBiggestContours, imgWarpColored, imgThresh],
                      [imgResult, imgRawDrawing, imgInvWarp, imgFinal])

    except:
        imgBlank = np.zeros_like(img)
        imageArray = ([img, imgGray, imgBlur, imgCanny],
                      [imgBlank, imgBlank, imgBlank, imgBlank],
                      [imgBlank, imgBlank, imgBlank, imgBlank])

    lables = [["Original", "Gray", "Blur", "Edges "],
              ["Contours", "Biggest Con", "Warp", "Threshold"],
              ["Result", "Raw Drawing", "Inv Warp", "Final"]]
    imgStacked = utlis.stackImages(imageArray, 0.3, lables)

    cv2.imshow("Final Result", imgFinal)
    cv2.imshow("Stacked Images", imgStacked)
    key = cv2.waitKey(1) & 0xFF  # Adjust the delay time
    if key == ord('s'):  # Press 's' to save the final image
        cv2.imwrite("FinalResult.jpg", imgFinal)
        name = input("Enter your name: ")
        email = input("Enter your email: ")
        save_to_excel(name, email, score)
        print("Marks saved to Excel.")

    if key == 27:  # Press 'Esc' to exit the loop
        break

# Release the capture and close the windows
cap.release()
cv2.destroyAllWindows()








