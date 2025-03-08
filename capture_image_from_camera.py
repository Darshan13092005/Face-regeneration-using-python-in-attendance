import cv2

cam_port = 0
cam = cv2.VideoCapture(cam_port)

if not cam.isOpened():
    print("Error: Camera not detected!")
    exit()

inp = input("Enter person's name: ")

while True:
    result, image = cam.read()

    if result:
        cv2.imshow('Preview', image)

        key = cv2.waitKey(1) & 0xFF

        if key == ord('s'):
            filename = inp + ".png"
            cv2.imwrite(filename, image)
            print(f"Image saved as {filename}")
            break
        elif key == ord('q'):
            print("Exiting without saving.")
            break
    else:
        print("No image detected. Please try again.")
        break

cam.release()
cv2.destroyAllWindows()
