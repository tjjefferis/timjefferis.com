---
layout: post
title: "Automating getting dates into an Outlook calendar"
date: 2019-04-29
blogger_url: "http://tjjteachersblog.blogspot.com/2019/04/automating-getting-dates-into-an-outlook-calendar.html"
---

If you are responsible, like me, for getting a large number of dates into an Outlook calendar you'll want some way of automating the process - particularly if the dates don't repeat regularly and there are <i>loads</i> of them!<br/>
<br/>
The first step of the process is to get all the dates into an Excel spreadsheet. The column headings you'll need if you use my method, from the left, are:<br/>
<br/>
<ul>
<li>Subject</li>
<li>Location</li>
<li>Required Attendees</li>
<li>Categories</li>
<li>Start Date</li>
<li>End Date</li>
<li>Start Time</li>
<li>End Time</li>
<li>Reminder</li>
<li>Duration</li>
<li>Optional Attendees</li>
<li>Description</li>
</ul>
<br/>
At this stage, you could just save your spreadsheet as a .csv file and use File &gt; Import to get the dates into the right calendar in Outlook. But if you do this you will quickly encounter a problem - the required attendee field won't be populated and so you won't get to share you imported dates with those you want to invite. Don't ask me why this failure occurs, but it does. The forums are largely silent on how to fix it :-(<br/>
<br/>
So you'll need to run a VBA script to get your events and attendees across into Outlook. Here are the steps:<br/>
<br/>
<br/>
<ul>
<li>Open Outlook</li>
<li>Go to File &gt; Options &gt; Customize Ribbon and make sure the Developer checkbox is ticked:</li>
</ul>
<br/>
<a href="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEjn4x-c5dbtlk9jDqKdgLo6oPCbIhfKcWzg1pilsvpy674CPd64I62f-Gsizyu9IDxqf_8umm3HMoUHrwt9fw2fLwFX8jPqbNU4i_95NGakl5LqPh48AOJXXC4AEu71-JiLb9tiEmgR4Fw/s1600/developer.PNG" imageanchor="1"><img border="0" data-original-height="360" data-original-width="364" height="316" src="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEjn4x-c5dbtlk9jDqKdgLo6oPCbIhfKcWzg1pilsvpy674CPd64I62f-Gsizyu9IDxqf_8umm3HMoUHrwt9fw2fLwFX8jPqbNU4i_95NGakl5LqPh48AOJXXC4AEu71-JiLb9tiEmgR4Fw/s320/developer.PNG" width="320"/></a><br/>
<br/>
<ul>
<li>Close the window.</li>
<li>Then click on Developer &gt; Visual Basic:</li>
</ul>
<br/>

<a href="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEi77F8Tmg5WbjqvhPZK_VO1CWTT6Hmy1KzPdZjGVeA6YqBGmv47Ag_vKJTSYwBdFzbEVoUDJgTwSFaVXQxmZojp9OhVDs70QQt4RVg6hqizUjV2NvBYNnXL8UIdXhf76XgFCOXf9Oga3Wk/s1600/vba.PNG" imageanchor="1"><img border="0" data-original-height="167" data-original-width="539" height="99" src="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEi77F8Tmg5WbjqvhPZK_VO1CWTT6Hmy1KzPdZjGVeA6YqBGmv47Ag_vKJTSYwBdFzbEVoUDJgTwSFaVXQxmZojp9OhVDs70QQt4RVg6hqizUjV2NvBYNnXL8UIdXhf76XgFCOXf9Oga3Wk/s320/vba.PNG" width="320"/></a>
<br/>
<br/>
<ul>
<li>Make sure you are in a project associated with 'ThisOutlookSession':</li>
</ul>
<br/>

<a href="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEh6I5ewhREdAd3ZKKTNRcxOPpaYd-IN9nR3dhKM1DFcF37P10qlGMLT1Rk-W4C1jIocqryVyWQtN5dpf0NJR7Ywwf4REdaQA7EERW95R7DfKxYHB4BFbFOPjEk_pQuAHI-anV5IEw3r0t4/s1600/vba1.PNG" imageanchor="1"><img border="0" data-original-height="174" data-original-width="321" height="173" src="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEh6I5ewhREdAd3ZKKTNRcxOPpaYd-IN9nR3dhKM1DFcF37P10qlGMLT1Rk-W4C1jIocqryVyWQtN5dpf0NJR7Ywwf4REdaQA7EERW95R7DfKxYHB4BFbFOPjEk_pQuAHI-anV5IEw3r0t4/s320/vba1.PNG" width="320"/></a>
<br/>
<br/>
<br/>
<ul>
<li>Now copy and paste this code into the VBA window (full disclosure - I didn't write this myself, I simply tweaked code very generously shared by Slipstick Systems <a href="https://www.slipstick.com/developer/code-samples/create-meetings-csv-file/" target="_blank">here</a>):</li>
</ul>
<blockquote class="tr_bq">
<i>Sub CreateMeetingsfromCSV()</i><i>    </i><i> ' Worksheet format: Subject, Location, Required Invitees, Categories, Start_Date, End_Date, Start_Time, End_Time, Reminder, Duration, Optional Attendees, Resource</i><i> ' Possible Values for Reminder Field is :'No Reminder','0 Minutes','1 Day','2 Days', '1 Week'</i><i>    </i><i>    Dim xlApp As Object 'Excel.Application</i><i>    Dim xlWkb As Object ' As Workbook</i><i>    Dim xlSht As Object ' As Worksheet</i><i>    Dim rng As Object 'Range</i><i>    Dim objAppt As Outlook.AppointmentItem</i><i>    Dim myAttendee As Outlook.Recipient</i><i>    Dim myOptional As Outlook.Recipient</i><i>    Dim myResource As Outlook.Recipient</i><i>    </i><i>    'Set xlApp = New Excel.Application</i><i>    Set xlApp = CreateObject("Excel.Application")</i><i>    </i><i>    strFilepath = xlApp.GetOpenFilename</i><i>    If strFilepath = False Then</i><i>        xlApp.Quit</i><i>        Set xlApp = Nothing</i><i>        Exit Sub</i><i>    End If</i><i>     </i><i>    Set xlWkb = xlApp.Workbooks.Open(strFilepath)</i><i>    Set xlSht = xlWkb.Worksheets(1)</i><i>    Dim iRow As Integer</i><i>    Dim iCol As Integer</i><i>    </i><i>    iRow = 2</i><i>    iCol = 1</i><i>     </i><i>    While xlSht.Cells(iRow, 1) &lt;&gt; ""</i><i>    </i><i>Set objAppt = Application.CreateItem(olAppointmentItem)</i><i>    </i><i>    Set myAttendee = objAppt.Recipients.Add(xlSht.Cells(iRow, 3))</i><i>          myAttendee.Type = olRequired</i><i>    'Set myOptional = objAppt.Recipients.Add(xlSht.Cells(iRow, 11))</i><i>          'myOptional.Type = olOptional</i><i>    'Set myResource = objAppt.Recipients.Add(xlSht.Cells(iRow, 12))</i><i>          'myResource.Type = olResource</i><i><br/></i><i>        With objAppt</i><i>                .Subject = xlSht.Cells(iRow, 1)</i><i>                .Location = xlSht.Cells(iRow, 2)</i><i>                .Categories = xlSht.Cells(iRow, 4)</i><i>                .Start = xlSht.Cells(iRow, 5) + xlSht.Cells(iRow, 7)</i><i>           ' Use either .Duration or .End</i><i>                .End = xlSht.Cells(iRow, 6) + xlSht.Cells(iRow, 8)</i><i>                .Body = xlSht.Cells(iRow, 12) &amp; vbNewLine &amp; vbNewLine &amp; "Dates fed from this sheet: http://gg.gg/dra28"</i><i>                '.Duration = xlSht.Cells(iRow, 10)</i><i>           ' This tells Outlook it's a meeting</i><i>               .MeetingStatus = olMeeting</i><i><br/></i><i>   Select Case xlSht.Cells(iRow, 9)</i><i>        Case "No Reminder"</i><i>            .ReminderSet = False</i><i>        Case "0 minutes"</i><i>            .ReminderSet = True</i><i>            .ReminderMinutesBeforeStart = 0</i><i>        Case "1 day"</i><i>            .ReminderSet = True</i><i>            .ReminderMinutesBeforeStart = 1440</i><i>        Case "2 days"</i><i>            .ReminderSet = True</i><i>            .ReminderMinutesBeforeStart = 2880</i><i>        Case "1 week"</i><i>            .ReminderSet = True</i><i>            .ReminderMinutesBeforeStart = 10080</i><i>    End Select</i><i>    </i><i>    For Each myAttendee In .Recipients</i><i>        myAttendee.Resolve</i><i>    Next</i><i>        .Save</i><i>        .Display</i><i>        '.Send ' hit the send button yourself to avoid Select names dialog</i><i>        End With</i><i>        iRow = iRow + 1</i><i>    Wend</i><i>    </i><i>    xlWkb.Close</i><i>    xlApp.Quit</i><i>    Set xlWkb = Nothing</i><i>    Set xlApp = Nothing</i><i> </i><i>End Sub</i></blockquote>
<ul>
<li>Once you are done click the green run button on the VBA taskbar.</li>
<li>You will be prompted to navigate to the Excel file you wish to import. Navigate to it.</li>
<li>Watch the magic happen!</li>
</ul>
<br/>
Note that this is a low-risk strategy because .Send is commented out. You will need to manually press send on each invitation - giving you the opportunity to sense check things one last time before they go out. If things have gone wrong you can delete the entries before anything bad happens!<br/>
<br/>
You'll want to edit the .Body contents too, of course.<br/>
<br/>
Have a play!<br/>
<br/>
<br/>