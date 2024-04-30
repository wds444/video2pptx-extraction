# video2pptx-extraction
A Python script that uses cv2 to extract slides shown in a recorded presentation.

Tested with mp4 files.

Edit the SLIDE_THRESHOLD to define how much your frames should differ from eachother.
Edit CONTENT_TRHESHOLD to define whether additional content being revealed on a slide should constitute a new slide or not.

Optionally define a Region of Interest where (not) to look, for example an overlayed presentor you want to ignore for slide update preferences.
Aspect ratio currently hardcoded to Widescreen.


