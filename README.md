# WinSum

WinSum is a simple, no-frills, script that allows you to get the checksum of a file and optionally compare it to a known checksum with just a couple of clicks of a mouse.

WinSum is written in VBscript and utilizes Windows' built-in CertUtil program.  Because WinSum relies on what is already in Windows, there are no additional programs that need to be installed.  This makes WinSum a very lightweight and portable utility.  WinSum currently provides MD5, SHA1, SHA256, and SHA512 checksums.


### To setup WinSum:

Setting up WinSum is easy.  Nothing to install, but there are some shortcuts that will need to be created to use WinSum.

1. Download and extract the WinSum files to your desired location.

1. Run `Create_SendTo_Shortcuts.vbs`.  This will create the needed shortcuts in your 'Send To' folder that gives you right click access to WinSum.

### To use WinSum:

Using WinSum is as simple as right-clicking your mouse.

1. Find a file that you would like to get the checksum.

1. Right-click on the file, go to 'Send To', then choose the algorithm.

1. You will see a box pop up showing you the file that was just checked as well as the checksum.  You also have the opportunity to paste in a known checksum to compare.  Click **OK**.

	*Note: Because the VBscript Inputbox does not expand, the calculated checksum may be partially obscured.  This is normal and you will see the full checksum in the next box.*

1. If you entered a checksum to compare, you will see the result as to whether they match or whether they are different.

	If you left the compare field blank, then you will see the full checksum.

	*Note: You can hit CTRL-C to copy the contents of the box.*

### To remove WinSum

If you find that WinSum doesn't meet your needs, it is easy to remove.

1. Run `Remove_SendTo_Shortcuts.vbs` to remove the previously created shortcuts.

	**OR**

	If you are more experienced, you could choose to remove the shortcuts manually.

1. With the shortcuts removed, you can now delete the WinSum files and/or directory.

### License
WinSum is under the BSD 3-Clause "New" or "Revised" License.