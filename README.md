<div align="center">

## Disable webbrowser control script errors


</div>

### Description

Like many people, I've played around with the web browser control and vb. One of things that has always ticked me off is the fact that you can try to cancel popups by using cancel=true in the newwindow2 event. But many times you get a small window showing some kind of script error. I've even tried using ppdisp=webbrowser2.object where webbrowser2.visible=false. No matter what I try I still get the window with a Internet Explorer Script Error. I started to look at the webbrowser.silent property but every time I set it to true it doesn't work. Finally, I know why. You see every time a webbrowser control is run, it automatically sets webbrowser.silent=false. Setting it to true during development doesn't work. What you need to do is run a seperate control. I used a timer control in the form load event, to set the .silent property to true. I then set the timer.enabled=false. The silent property is now set and you won't have to worry about script messages again. So far every web site with popups has been blocked and no script errors. In the source example I also included a way to control popups using the ctrl key.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-03-10 14:21:02
**By**             |[Karl J\. Sak](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/karl-j-sak.md)
**Level**          |Beginner
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Disable\_we1557483102003\.zip](https://github.com/Planet-Source-Code/karl-j-sak-disable-webbrowser-control-script-errors__1-43907/archive/master.zip)








