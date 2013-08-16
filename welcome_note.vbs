Set Sapi = Wscript.CreateObject("Sapi.SpVoice")
Sapi.speak "Welcome "
if hour(time) < 12 then
Sapi.speak "Good Morning  "
else
if hour(time) > 12 & hour(time) <17 then
Sapi.speak "Good Afternoon "
else
if hour(time) > 17 & hour(time)<22 then
Sapi.speak "Good Evening "
else 
if hour(time)>22
Sapi.speak "Its late in the night . Wish you good night in advance "
end if
end if
end if
end if
Sapi.speak "System Information"
Sapi.speak "This is "
Sapi.speak weekdayname(weekday(date))
Sapi.speak " the current date."
Sapi.speak day(date)
Sapi.speak monthname(month(date))
Sapi.speak year(date)
Sapi.speak "System Login at "


if hour(time) > 12 then
Sapi.speak hour(time)-12
else
if hour(time) = 0 then
Sapi.speak "12"
else
Sapi.speak hour(time)
end if
end if


if minute(time) < 10 then
Sapi.speak "o"
if minute(time) < 1 then
Sapi.speak "clock"
else
Sapi.speak minute(time)
end if
else
Sapi.speak minute(time)
end if

if hour(time) > 12 then
Sapi.speak "P.M."
else
if hour(time) = 0 then
if minute(time) = 0 then
Sapi.speak "Midnight"
else
Sapi.speak "A.M."
end if
else
if hour(time) = 12 then
if minute(time) = 0 then
Sapi.speak "Noon"
else
Sapi.speak "P.M."
end if
else
Sapi.speak "A.M."
end if
end if
end if

Sapi.speak " Have a nice day."
