**EDM Mobile**

This is one in a series of programs written by Harold Dibble and Shannon McPherron to help archaeologists excavate archaeological sites.  In the late 1990s we realized that the hardware we were using then (DOS based portable computers) would soon be no longer available.  We then created two new version of the EDM program: one for Windows and one for what was then called Windows CE.  The latter ran on small handheld devices a bit like smartphones before smartphones existed.  The DOS program we were porting was written in QBASIC.  I don't remember what Microsoft called the programming environment for Windows CE, but we used a variant of BASIC.  Eventually Microsoft dropped CE and replaced it with the short-lived Windows Mobile and they introduced the .NET programming environment.  So I then ported EDM CE over to EDM Mobile, and we used this program at countless sites over the years running on devices like the Trimble Nomad (mine are still running after 20 years).

A few years ago, I rewrote EDM from scratch into Python (EDMpy), and I used a carefully selected set of libraries that would help me keep Android compatibility.  While EDMpy works well on Windows and Linux machines (and may work as well on MacOS), I have not yet solved some technical issues related to BlueTooth to be able to use it on an Android phone.  So while I really wanted people to stop using it, there is still a high demand for EDM Mobile for archaeologists who can't easily run a Windows laptop or tablet at their site.  And while Microsoft hasn't supported the operating system in many, many years, you can still find handheld devices based around it that will run this program.

In the mean time, meaning the last 25 years, Leica changed the communication format they use for communicating with total stations.  It is now called GeoCOM.  For a long time they made total stations that supported both formats, but in the last years they dropped the old format entirely.  This means that while you can still find computers to run EDM Mobile (or use old ones), you couldn't pair it with a new station.  Thus in April, 2025, I was persuaded to update the program (for the first time in maybe 6 years) to include support for newer Leica stations.  This version is called 1.10 and the station type is called Leica-Geo.

With this change I have decided to make the code base public.  It is actually something I intended to do a few years ago, but a) I needed to clean the code a bit more and b) it is hard to release software period but I think especially hard to release something that you wrote 25 years ago and that itself was a port of something much older.  When I look back at this code, there are definitely aspects that make me cringe.  But, and I think this is really important, so important that I really shouldn't bury it in the middle of a paragraph, I think researchers have a right to inspect, audit, and potentially modify the code they use to collect expensive and absolutely irreplaceable data.  It is true that with a total station there are ways to check that the data are correct and internally consistent (for instance, verification datums and plotting), but still the principle remains.  Research software should be open, maybe especially when it is written by non-professionals.

So here it is - EDM Mobile.  If you want to recreate my work environment, I use an Oracle VM VirtualBox running Windows XP, and in this I run Microsoft Visual Studio 2008.  Maybe now you understand why I wanted to be done with this program.

If you do not want the source code and only want the latest version of the program, then look in [edm_mobile/bin/release](https://github.com/surf3s/EDM_Mobile/tree/main/edm_mobile/bin/Release).  You will also need the three DLL files included there.  I suggest you read the page on the [OldStoneAge website](https://www.oldstoneage.com/osa/tech/edmmobile/) for how to install this program.

**Version Information**

Version 1.10

Added Leica-Geo for newer Leica total station that only support the GeoCOM format.

Version 1.9 

Unit defaults will only be changed when Suffix = 0.  This makes it much easier to take a shot, change its ID, change its suffix, and then carry on after this with what would have been normally the next number.  This is important for closing buckets.

Fixed annoying bug wherein you had to click away and then click back to first field in the edit list to actually get the menu to activate.

