# Binary metamorphosis

These VB6 applications are able to convert any executable file (ie. <kbd>.exe</kbd> or any type of file in fact) to VB6 source code which can be included back into a VB6 application. Once inserted into a new VB6 application, the source code can be used to completely restore the original executable file to disk in the same directory as the application that incorporates the code. The point here is that one may compile an executable file that contains another executable file inside. Once the new executable file is executed, it is able to write the embedded executable file on disk as an independent executable file. Note that there are three separate versions of this application in the repository, but the difference between them is very small and it is more related to the user interface than the source code. Unless a specific executable file is chosen by the user for conversion, by default, these applications will convert a <kbd>tini.exe</kbd> file to hexadecimal in order to demonstrate the method.

<img src="https://github.com/Gagniuc/Binary-metamorphosis/blob/main/img/1.png?raw=true" alt="Binary metamorphosis">

# Regarding tini.exe or tini.executable

Antivirus engines that lack sophistication or professionalism may popup up false detections in connection to <kbd>tini.exe</kbd>. It is true that the method shown here has been used in the malware world countless times to hide executable files. However, this method proves to be extremely useful for virtuous purposes, such as software automation or for removing dependencies by incorporating certain resources into the compiled file. Thus, coding methods should not be used as signatures for detection by antivirus engines. Otherwise, this would be similar to: "<i>we no longer use uranium in nuclear power plants because atomic bombs can be made with it</i>". Nonetheless, the VB6 source code of <kbd>tine.exe</kbd> is available in the <kbd>tini</kbd> folder and it can be compiled at will. An MD5 comparison of the newly compiled file and the old file uploaded here verifies that <kbd>tini.exe</kbd> is NOT malware. As one can see in the source code of <kbd>tini.exe</kbd>, the application is completely empty and contains a simple window with the message:

<img src="https://github.com/Gagniuc/Binary-metamorphosis/blob/main/img/tini.png" alt="tini.exe">

The following signatures are expected from the file you are compiling or from the <kbd>tini.exe</kbd> files (aka <kbd>tini.executable</kbd>) already compiled and uploaded here:

```
MD5 File Checksum: f0b950acbf1af90eb8e9a52cc6799a08
SHA512 File Hash:  9a5c1e2cb40b1a0777e9975c52fdc6798b27e30a11090da6107a63b265220caa53d7301c10920d91033c05233bd25c65c216a1ca274eb98e429c63de453ec394
```

Note that tini.exe is a small file (16Kb) and it is used here for providing an exemplification by using smaller VB6 source code. However, one may use any executable file. Moreover, this method can embed any type of file, not only executables.


# Hexadecimal (or hex)

The VB6 application shown here uses the hexadecimal system to encode the binary content of an executable file. But hat is hexadecimal? Hexadecimal is a base 16 system that is used for binary representation. A hex digit can be any of the following 16 digits: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C, D, E, F, were each hex digit represents a 4-bit binary sequence. The table below shows the association of each hex digit with the equivalent values in binary and denary:

| Denary | Binary | Hexadecimal |
| :---   |  :---: |     ---:    |
| 0      | 0000   | 0           |
| 1      | 0001   | 1           |
| 2      | 0010   | 2           |
| 3      | 0011   | 3           |
| 4      | 0100   | 4           |
| 5      | 0101   | 5           |
| 6      | 0110   | 6           |
| 7      | 0111   | 7           |
| 8      | 1000   | 8           |
| 9      | 1001   | 9           |
| 10     | 1010   | A           |
| 11     | 1011   | B           |
| 12     | 1100   | C           |
| 13     | 1101   | D           |
| 14     | 1110   | E           |
| 15     | 1111   | F           |


It is much less time consuming to write numbers as hex than to write them as binary numbers. An 8-bit binary number can be written using only two different hex digits, one hex digit for each group of 4-bits. An 16-bit binary number can be written using only four different hex digits, and so on. For example, an 8-bit binary number like <kbd>10100101</kbd> would be <kbd>A5</kbd> in hex. An 16-bit binary number like <kbd>1011110001101110</kbd> would be <kbd>BC6E</kbd> in hex. To continue with the exemplification, a hex representation like <kbd>CCCCC7</kbd> would be <kbd>110011001100110011000111</kbd> in binary.
