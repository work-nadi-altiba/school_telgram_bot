# importing the module
import cv2
import sys
# function to display the coordinates of
# of the points clicked on the image
def click_event(event, x, y, flags, params):
    # y-=297
    # y = abs(y)
     # checking for left mouse clicks
    if event == cv2.EVENT_LBUTTONDOWN:
        # displaying the coordinates
        # on the Shell
        print('[' , str((x/2)-7), ',', str((abs(y-297*2)/2)-7-17),']' ,sep='',end=',' )
        # displaying the coordinates
        # on the image window
        font = cv2.FONT_HERSHEY_SIMPLEX
        cv2.putText(img, '+'+ str((x/2)-7) + ',' +
                    str((abs(y-297*2)/2)-7-17) , (x-7,y+4) , font ,
                    0.5, (255, 0, 0), 2 )
        cv2.imshow('image', img)
    # checking for right mouse clicks 
    
def image_resize(image, width = None, height = None, inter = cv2.INTER_AREA):
    # initialize the dimensions of the image to be resized and
    # grab the image size
    dim = None
    (h, w) = image.shape[:2]

    # if both the width and height are None, then return the
    # original image
    if width is None and height is None:
        return image

    # check to see if the width is None
    if width is None:
        # calculate the ratio of the height and construct the
        # dimensions
        r = height / float(h)
        dim = (int(w * r), height)

    # otherwise, the height is None
    else:
        # calculate the ratio of the width and construct the
        # dimensions
        r = width / float(w)
        dim = (width, int(h * r))

    # resize the image
    resized = cv2.resize(image, dim, interpolation = inter)

    # return the resized image
    return resized

# driver function
if __name__=="__main__":
    # reading the image
    img = cv2.imread('document-page1-1.jpg' ,1)
    width = 210*2
    height = 297*2
    dim = (width, height)
    
    img = image_resize(img , width= width , height = height)
    # resize image
    # img = cv2.resize(img,  (100, 50))

    # displaying the image
    cv2.imshow('image', img )
    # setting mouse handler for the image
    # and calling the click_event() function
    cv2.setMouseCallback('image', click_event)
    # wait for a key to be pressed to exit
    cv2.waitKey(0)
    # close the window
    cv2.destroyAllWindows()

# TODO: احذف كل التعليقات و خلي بس المهمة  والمفيدة 
# TODO: استخدم كنتر او كيو تي في البرنامج 
# TODO: احفظ الاحداثيات في ملف واحد و منظم باسم الصفحة و اسم للخانة 