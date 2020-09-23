<div align="center">

## Some Spells about IrDa and Tiny Devices \+ Embedded Visual Basic 3\.0 code\!


</div>

### Description

Some Spells and tips on how Develop Good Program for IrDa devices and solve some annoing connection problems between IrDa devices.

A rush of code written in

embedded visual basic 3.0 (Windows CE 3.0),

explain the facts.

See also my article on Generic IrDa systems!

here:

http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=4725&lngWId=3

PLEASE DOWNLAD THE ZIP BECAUSE PLANET SOURCE CODE DOES NOT DISPLAY CORRECTLY THE HTML!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-09-09 17:45:12
**By**             |[michele berardi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michele-berardi.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows CE](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-ce__1-41.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Some\_Spell1290859102002\.zip](https://github.com/Planet-Source-Code/michele-berardi-some-spells-about-irda-and-tiny-devices-embedded-visual-basic-3-0-code__1-38828/archive/master.zip)





### Source Code

```
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Some Spells on: IrDa (Infrared) devices programming</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<p>Some Spells on: &quot;IrDa (Infrared) devices programming&quot;<br>
 whit code snippet written in: <br>
 Embedded Visual Basic 3.0 (Windows Ce 3.0)<br>
 (C) 2002 Berardi Michele<br>
 http://web.tiscali.it/mberardi <br>
<p><br>
 The infrared devices are joined to the BCR (a serial device).<br>
 <br>
 The BCR cannot send off messages to the irda peripherals<br>
 about its speed of communication, therefore is necessary<br>
 to set up and initialize on both irda peripherals,<br>
 the same speed of communication.<br>
 <br>
 The values of default of habit they are: 9600, N,8,1.<br>
 <br>
 The device Irda need for running 12 milliampere<br>
 power suppyed from the BCR, either from a battery or<br>
 external power supply.<br>
 <br>
 The connection means habit throught a serial connector<br>
 9 pin female, which takes the power<br>
 from the handshaking's pin.<br>
 <br>
 To make one's will the communication Irda using a<br>
 Desktop PC provided of an Irda Eye.<br>
 You can use a simple terminal emulator<br>
 on the desktop setting the same parameters<br>
 as the pocket pc device ( 9600,N,8,1 ) <br>
 and obvious choosing the serial port<br>
 assigned to Irda device (the port must be<br>
 correctly installed on system).<br>
 <br>
 If it is wanted writeran applications for the IrDa com,<br>
 you are remembered to settings highs (true for the VB)<br>
 the values DTR and RTS in the property of the object COM,<br>
 otherwise the communication cannot happen<br>
 for scarcity of energy,<br>
 besides for the object COM in matter is necessary set up<br>
 the propertyes:<br>
 <br>
 Rthreshold = 1 and Strheshold = 0.<br>
 <br>
 Example of evb code:
<table cellpadding=5 width="100%" bgcolor=#e0e0e0 border=0>
 <tbody>
  <tr>
   <td> <p>'<br>
     ' how to create a com object and how it (described above)<br>
     ' assign a generic Irda port (as usual COM4)<br>
     '<br>
     <br>
     Option Explicit<br>
     <br>
     Private Sub Receive(data)<br>
     txtReceived.SelStart = Len (txtReceived.Text)<br>
     txtReceived.SelText = data<br>
     End Sub<br>
     <br>
     Private Sub InviaTesto_Click ()<br>
     Comm1.Output = txtReceived.Text<br>
     End Sub<br>
     <br>
     Private Sub Comm1_OnComm ()<br>
     If Comm1.CommEvent = 2 Then Call Receive(Comm1.Input)<br>
     End Sub<br>
     <br>
     Private Sub Form_Load ()<br>
     If Not Comm1.PortOpen Then Comm1.PortOpen = True<br>
     End Sub<br>
    </p>
    </td>
  </tr>
 </tbody>
</table>
 <br>
<br>
Berardi Michele<br>
 Senior Developer<br>
 &quot;customize your opportunities!&quot;<br>
 Mobile: + 39 347 319 2000<br>
 E-Mail(S) :<br>
mfxaub@tin.it<br>
 03473192000@vizzavi. it<br>
 Web:<br>
 http://web.tiscalinet.it/mberardi<br>
 <br>
<p></p>
<p><br>
</p>
<p>&nbsp; </p>
</body>
</html>
```

