<%@ Import Namespace=System.Drawing %>
<%@ Import Namespace=System %>
<%@ Import Namespace=System.Web %>

' Properly Scaled Thumbnails
' V 0.01

' Modified code from:
' The Code Project
' http://www.codeproject.com/aspnet/Thumbnail_Image.asp?print=true
' 
' Modified by glypher
' Contact: glypher@glypher.com
'
' this file can be found at:
' http://www.afn.org/~xonix/downloads/software/thumbnailimage.zip
' 
' this is being used at www.gainesvillebands.com on the flyers page.
'





<html>
<script language="VB" runat="server">

Sub Page_Load(Sender As Object, E As EventArgs)
	
' Dim all Variables
	Dim OrginalImg, Thumb As System.Drawing.Image
	Dim Rootpath, FileName As String
	Dim imgHeight, imgWidth, maxWidth, MaxHeight As Integer
	Dim inp As New IntPtr()
 
' Get Root Application Folder
	rootpath = Server.MapPath("/")

' Get filename
	FileName = rootpath & Request.QueryString("FileName")

' Attempt to populate the original image object	
	Try
		orginalimg = orginalimg.FromFile(FileName)

' Get the current Height and width of the image
		imgHeight = orginalimg.Height
		imgWidth = orginalimg.Width

' Get width using QueryString.
		If Request.QueryString("width") = Nothing Then
			maxWidth = orginalimg.Width
		ElseIf Request.QueryString("width") = 0 Then
			maxWidth = 100
		Else
	 		maxWidth = Request.QueryString("width")
		End If
 
' Get height using QueryString.
		If Request.QueryString("height") = Nothing Then
			MaxHeight = orginalimg.Height
		ElseIf Request.QueryString("height") = 0 Then
			MaxHeight = 100
		Else
			MaxHeight = Request.QueryString("height")
		End If

' Check to see if the image even needs scaled
		If imgWidth > maxWidth Or imgHeight > MaxHeight Then

' Determine what dimension is off by more
			Dim deltaWidth As Integer = imgWidth - maxWidth
			Dim deltaHeight As Integer = imgHeight - MaxHeight
			Dim scaleFactor As Double
 
			If deltaHeight > deltaWidth Then

' Use the Height to set the scale factor
				scaleFactor = MaxHeight / imgHeight
			Else

' Use the Width to set the scale factor
				scaleFactor = maxWidth / imgWidth
			End If
 
' Set the new Scaled the image Size 
			imgWidth *= scaleFactor
			imgHeight *= scaleFactor
		End If

' If the population fails get error.gif (must exist or will cause an error)
	Catch
		orginalimg = orginalimg.FromFile(rootpath & "error.gif")
	End Try

' set the thumbnail width and height by the new correct scale
	thumb = orginalimg.GetThumbnailImage(imgWidth, imgHeight, Nothing, inp)
 
' Sending JPG Response to the browser. 
	Response.ContentType = "image/jpeg"
	thumb.Save(Response.OutputStream, Imaging.ImageFormat.Jpeg)
 
' Disposing the objects.
	orginalimg.Dispose()
	thumb.Dispose() 

  End Sub
</script>
</html>
