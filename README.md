# The Student Relocation Script  
A script that organises students into groups based on their preferences  
  
## Libraries Required  
To install all required libraries, run  
`pip -r requirements.txt`  
  
## Usage  
`python Student_Relocation_Script.py`  
  
Terminology:\
DS - dissaproval score. This is how much the student and other students dislike the student being located there. The LOWER the better.\
Kick - move a student from a group into a group with only them in it\
\
This script takes data from a excel spreadsheet that is exported from a microsoft form.\
The form's fist question should be required, the other should not be.\
The form's format should be this:\
\
What is your name? Answer in this format: First name [space] last name (Example: John Doe)\
Who do you want to sit with? Answer in the same format as the previous question.\
Who else do you want to sit with? Answer in the same format as the previous question.\
Who else do you want to sit with? Answer in the same format as the previous question.\
Who else do you want to sit with? Answer in the same format as the previous question.
