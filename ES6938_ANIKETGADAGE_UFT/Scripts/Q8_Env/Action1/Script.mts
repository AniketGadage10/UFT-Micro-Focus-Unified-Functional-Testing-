'Question 8
'Add 5 Environment variables for test and change value of them run time . Display changed values

print "Before Update : -"&Environment.Value("UserId")

Environment.Value("UserId")="Aniket1018"

print "After Update : -"&Environment.Value("UserId")
