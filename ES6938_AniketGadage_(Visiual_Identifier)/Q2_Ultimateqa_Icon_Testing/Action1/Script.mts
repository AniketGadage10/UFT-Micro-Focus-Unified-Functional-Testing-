'write program to open https://ultimateqa.com/automation website and verify images of Facebook, Tweeter, Linked-In icon using insite objects.

Option Explicit

SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://ultimateqa.com/automation"

If Browser("Brw_UltimateQA").InsightObject("Facebook_Icon").Exist(4) THEN
	PRINT "Facebook Icon Is Present "
End If
If  Browser("Brw_UltimateQA").InsightObject("Linkedin_Icon").Exist(4)  THEN
	PRINT "Linkedin Icon Is Present "
End If
If Browser("Brw_UltimateQA").InsightObject("Instagram_icon").Exist(4) Then
	PRINT "Instagram Icon Is Present "
End If
