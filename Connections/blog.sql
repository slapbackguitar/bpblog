USE [Blog]
GO
/****** Object:  Table [dbo].[tblAuthor]    Script Date: 11/07/2006 10:51:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblAuthor](
	[fldAuthorID] [int] IDENTITY(1,1) NOT NULL,
	[fldAuthorUsername] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldAuthorRealName] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldAuthorEmail] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldAuthorWebsite] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldAuthorBlurb] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldAuthorPassword] [nvarchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Approved] [smallint] NULL DEFAULT ((0)),
	[fldAdmin] [smallint] NULL DEFAULT ((0)),
 CONSTRAINT [aaaaatblAuthor_PK] PRIMARY KEY NONCLUSTERED 
(
	[fldAuthorID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblBlog]    Script Date: 11/07/2006 10:52:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblBlog](
	[BlogID] [int] IDENTITY(1,1) NOT NULL,
	[BlogHeadline] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[BlogHTML] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BlogDate] [datetime] NOT NULL DEFAULT (getdate()),
	[BlogCat] [int] NULL DEFAULT ((0)),
	[BlogAuthor] [int] NULL DEFAULT ((0)),
	[BlogCommentInclude] [smallint] NULL DEFAULT ((1)),
	[BlogReadMore] [smallint] NULL DEFAULT ((0)),
	[BlogDraft] [smallint] NULL DEFAULT ((0)),
 CONSTRAINT [aaaaatblBlog_PK] PRIMARY KEY NONCLUSTERED 
(
	[BlogID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblBlogRSS]    Script Date: 11/07/2006 10:52:21 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblBlogRSS](
	[rssID] [int] IDENTITY(1,1) NOT NULL,
	[blogTitle] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[blogSubTitle] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[blogDesc] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[blogURL] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[blogImage] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[blogAuthor] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[blogEmail] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[blogPosts] [int] NULL DEFAULT ((20)),
	[blogLayout] [int] NOT NULL DEFAULT ((1)),
 CONSTRAINT [aaaaatblBlogRSS_PK] PRIMARY KEY NONCLUSTERED 
(
	[rssID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblCat]    Script Date: 11/07/2006 10:52:56 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblCat](
	[CatID] [int] IDENTITY(1,1) NOT NULL,
	[CatName] [nvarchar](200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CatDesc] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [aaaaatblCat_PK] PRIMARY KEY NONCLUSTERED 
(
	[CatID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblComment]    Script Date: 11/07/2006 10:53:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblComment](
	[commentID] [int] IDENTITY(1,1) NOT NULL,
	[blogID] [int] NOT NULL DEFAULT ((0)),
	[commentDate] [datetime] NOT NULL DEFAULT (CONVERT([datetime],CONVERT([varchar],getdate(),(1)),(1))),
	[commentName] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[commentEmail] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[commentURL] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[commentHTML] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[commentInclude] [int] NOT NULL DEFAULT ((0)),
 CONSTRAINT [aaaaatblComment_PK] PRIMARY KEY NONCLUSTERED 
(
	[commentID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblGallery]    Script Date: 11/07/2006 10:53:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblGallery](
	[fldGalleryID] [int] IDENTITY(1,1) NOT NULL,
	[fldGalleryTitle] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldGalleryDesc] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldGalleryPic] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[fldGalleryCreated] [datetime] NULL DEFAULT (getdate()),
	[fldGalleryUser] [int] NOT NULL DEFAULT ((1)),
 CONSTRAINT [aaaaatblGallery_PK] PRIMARY KEY NONCLUSTERED 
(
	[fldGalleryID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO



USE [Blog]
GO
/****** Object:  Table [dbo].[tblGalleryConfig]    Script Date: 11/07/2006 10:53:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblGalleryConfig](
	[fldGalleryConfigID] [int] IDENTITY(1,1) NOT NULL,
	[fldGalleryTitleThumb] [smallint] NULL DEFAULT ((0)),
	[fldGalleryThumb] [smallint] NULL DEFAULT ((0)),
 CONSTRAINT [aaaaatblGalleryConfig_PK] PRIMARY KEY NONCLUSTERED 
(
	[fldGalleryConfigID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblLayout]    Script Date: 11/07/2006 10:54:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblLayout](
	[layoutid] [int] IDENTITY(1,1) NOT NULL,
	[layout1] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[layout2] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[layout3] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[layout4] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[layout5] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[layoutTitle] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
 CONSTRAINT [aaaaatblLayout_PK] PRIMARY KEY NONCLUSTERED 
(
	[layoutid] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

USE [Blog]
GO
/****** Object:  Table [dbo].[tblPage]    Script Date: 11/07/2006 10:54:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblPage](
	[PageName] [nvarchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[PageTitle] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[PageHTML] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PageDate] [datetime] NOT NULL DEFAULT (getdate()),
 CONSTRAINT [aaaaatblPage_PK] PRIMARY KEY NONCLUSTERED 
(
	[PageName] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

INSERT INTO [blog].[dbo].[tblAuthor]
           ([fldAuthorUsername]
           ,[fldAuthorRealName]
           ,[fldAuthorEmail]
           ,[fldAuthorWebsite]
           ,[fldAuthorBlurb]
           ,[fldAuthorPassword])
     VALUES
           ('admin'
           ,'asdf'
           ,'asdf@asdf.com'
           ,'http://domain.com'
           ,'Test'
           ,'password')
           
INSERT INTO [blog].[dbo].[tblBlog]
           ([BlogHeadline]
           ,[BlogHTML]
           ,[BlogDate]
           ,[BlogCat]
           ,[BlogAuthor]
           ,[BlogCommentInclude])
     VALUES
           ('Introduction'
           ,'<BLOCKQUOTE dir=ltr style="MARGIN-RIGHT: 0px"><P><EM>"All happy families are alike; every unhappy family is unhappy in its own way." - Leo Tolstoy</EM></P></BLOCKQUOTE>'
           ,'1/9/2004 9:01:19 AM'
           ,1
           ,1
           ,1)           
           
INSERT INTO [blog].[dbo].[tblBlogRSS]
           ([blogTitle]
           ,[blogSubTitle]
           ,[blogDesc]
           ,[blogURL]
           ,[blogImage]
           ,[blogAuthor]
           ,[blogEmail]
           ,[blogPosts])
     VALUES
           ('blog title'
           ,'sub title'
           ,'description'
           ,'http://domain.com/blog/'
           ,'http://domain.com/blog/blog_button.jpg'
           ,'asdf'
           ,'asdf@asdf.com'
           ,20)
           
INSERT INTO [blog].[dbo].[tblCat]
           ([CatName]
           ,[CatDesc])
     VALUES
           ('Blog'
           ,'Description')   
           
INSERT INTO [blog].[dbo].[tblPage]
           ([PageName]
           ,[PageTitle]
           ,[PageHTML]
           ,[PageDate])
     VALUES
           ('about'
           ,'About BP Blog'
           ,'<P>bp blog is a blog software coded in ASP.</P>'
           ,'1/9/2004 9:01:19 AM') 
           
INSERT INTO [blog].[dbo].[tblPage]
           ([PageName]
           ,[PageTitle]
           ,[PageHTML]
           ,[PageDate])
     VALUES
           ('thanks'
           ,'Thank you for your comments!'
           ,'Your comments are subject to approval.'
           ,'2/15/2005 1:27:39 PM')  
           
INSERT INTO [blog].[dbo].[tblPage]
           ([PageName]
           ,[PageTitle]
           ,[PageHTML]
           ,[PageDate])
     VALUES
           ('thankyou'
           ,'Thank you for registering!'
           ,'Thank you for registering.'
           ,'2/15/2005 1:27:39 PM')                
           
INSERT INTO [blog].[dbo].[tblGalleryConfig]
           ([fldGalleryTitleThumb]
           ,[fldGalleryThumb])
     VALUES
           (200
           ,100) 
           
INSERT INTO [blog].[dbo].[tblLayout]
           ([layout1]
           ,[layout2]
           ,[layout3]
           ,[layout4]
           ,[layout5]
           ,[layoutTitle])
     VALUES
           ('asdf'
           ,'asdf'
           ,'asdf'
           ,'asdf'
           ,'asdf'
           ,'Default')             
           
/*
Insert this for the default layout via the admin screen

Title: Black Minima
Layout1:
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="generator" content="BP Blog 7.0" />
<link rel="copyright" href="http://creativecommons.org/licenses/by-nc-sa/2.5/" />
Layout2:
<link href="styles-site.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<script src="js/prototype.js" type="text/javascript"></script>
<script src="js/scriptaculous.js?load=effects" type="text/javascript"></script>
	<script src="js/lightbox.js" type="text/javascript"></script>
</head>
<BODY>
<DIV id=container>
<DIV id=header>
Layout3:
</DIV>
<DIV id=content>
<DIV id=main>
<!-- AdSense Code -->
<style type="text/css">
<!--
.adsensefloat {
	margin: 4px;
	padding: 2px;
	float: right;
	margin-top:10px;
}
-->
</style>
<div class="adsensefloat">
<script type="text/javascript"><!--
google_ad_client = "pub-0172387945911790";
google_ad_width = 120;
google_ad_height = 240;
google_ad_format = "120x240_as";
google_ad_channel ="7792557859";
google_color_border = "000000";
google_color_bg = "000000";
google_color_link = "AADD99";
google_color_text = "CCCCCC";
google_color_url = "AADD99";
//--></script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>
</div>
<!-- AdSense Code -->
Layout4:
</DIV>
<DIV id=sidebar>
<h2 class=sidebar-title>Menu</h2>
<ul>
<li><a href="default.asp">Home</a></li>
<li><a href="template.asp?pagename=about">About BP Blog</a></li>
<li><a href="template_gallery.asp">Gallery</a></li>
</ul>
Layout5:
</DIV>
 </DIV>
<DIV id=footer><P>Powered by <a href="http://blog.betaparticle.com" title="Powered by BP Blog 7.0">BP Blog 7.0</a>
		 | <a href="rss.xml">Feed (RSS)</a></P>
</DIV>	
</DIV>	 
</BODY>
</HTML>
*/
             