import win32com.client as wincl
import cv2
import numpy as np
import threading





show_Boxes = False

face_cascade = cv2.CascadeClassifier("haarcascade_frontalface_default.xml")

eye_cascade = cv2.CascadeClassifier("haarcascade_eye.xml")

smile_cascade = cv2.CascadeClassifier("Lib/site-packages/cv2/data/haarcascade_smile.xml")

profile_cascade = cv2.CascadeClassifier("Lib/site-packages/cv2/data/haarcascade_profileface.xml")

cap = cv2.VideoCapture(0)

speak = wincl.Dispatch("SAPI.SpVoice")

choices_ger = ["du spaßt", "eigentlich ein hässliches lächeln", "kannst eigentlich gleich sterben",
               "wieso überhaupt lächeln mit so nem behindertem lächeln", "du hurensohn",
               "Wieso grinst du so behindert", "also gut sieht das nicht aus", "einem Affen würdest du gefallen"]
               
choices_en = ["You ugly piece of cardboard lol", "you wanker", "a smile that haunts me like hell",
              "dont make me watch this again", "come down to earth again, no need to smile", "forcing a smile is not gonna help you", "better turn the camera off next time",
              "mate, you're scaring the kids", "woah, yeah, enough of that", "don't think this is even slightly funny", "yeah, better hide the pain"]



def curse():
    
    i = np.random.randint(0, len(choices_en))
    speak.Speak(choices_en[i])



def cv():
    while True:
        ret, img = cap.read()
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # smiles = smile_cascade.detectMultiScale(gray, 1.8, 20)

        faces = face_cascade.detectMultiScale(gray, 1.1, 4)
        profiles = profile_cascade.detectMultiScale(gray, 1.1, 4)
        for (x, y, w, h) in faces:
            if show_Boxes:
                cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 2)

            roi_gray = gray[y:y + h, x:x + h]
            roi_color = img[y:y + h, x:x + h]
            eyes = smile_cascade.detectMultiScale(gray, 2, 20)
            smiles = smile_cascade.detectMultiScale(roi_gray, scaleFactor=1.8, minNeighbors=20)

            for (sx, sy, sw, sh) in smiles:

                cv2.rectangle(roi_color, (sx, sy), (sx + sw, sy + sh), (0, 255, 255), 2)

                if threading.active_count() < 3: #  start cursing when smile detected and no curse going already
                    c = threading.Thread(target=curse)
                    c.start()

            for (ex, ey, ew, eh) in eyes:

                if show_Boxes:
                    cv2.rectangle(roi_color, (ex, ey), (ex + ew, ey + eh), (0, 255, 0), 2)
        for (x, y, w, h) in profiles:
            if show_Boxes:
                cv2.rectangle(img, (x, y), (x + w, y + h), (255, 255, 0), 2)

        cv2.imshow("img", img)
        k = cv2.waitKey(30) & 0xff
        if k == 27:
            break


t = threading.Thread(target=cv)
t.start()
t.join()

cap.release()
cv2.destroyAllWindows()
