<div align="center">

## Take Control of the Compiler


</div>

### Description

Access CPU registers, write true in-line C, C++, and assembly, hook API calls made by other programs, export your functions to a non-ActiveX DLL (in other words: make APIs), call functions by address, etc, etc, etc. The potential is mind boggling!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-07-09 13:13:36
**By**             |[Jonathan Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-smith.md)
**Level**          |Advanced
**User Rating**    |4.7 (66 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Take Contr22451792001\.zip](https://github.com/Planet-Source-Code/jonathan-smith-take-control-of-the-compiler__1-24616/archive/master.zip)





### Source Code

<p align="center"><b><font face="Verdana" size="6" color="#0000FF">Take Control
of the Compiler<br>
</font><i><font face="Verdana" color="#0000FF" size="4">For VB5 and VB6</font></i></b></p>
<p align="center"> </p>
<p align="left"><i>Author's Note: This is article is a rewritten excerpt of an
original written by John Chamberlain, a director of software development at
Clinical NetwoRx (cnrx.com). He can be reached by e-mail at <a href="mailto:jchamber@lynx.dac.neu.edu">jchamber@lynx.dac.neu.edu</a>.
Give credit and props for the original code and article to him. I am merely
rewriting this to put everything into a better perspective for most of the
people on PSC.</i></p>
<p align="left"><b>Objectives</b></p>
<p align="left">In the accompanying article and source code, you will learn how
to write an add-in that allows you to do the following:</p>
<ol>
 <li>
 <p align="left">View your application's native/object source</li>
 <li>
 <p align="left">Perform selective compilation of your project</li>
 <li>
 <p align="left">Statically link non-VB modules (use <i>true</i><b> </b>in-line
 C, C++, and assembly code in your projects)</li>
 <li>
 <p align="left">Export functions in your program to a normal, non-ActiveX
 DLL (an API DLL)</li>
 <li>
 <p align="left">Hook API calls by patching the import address table (IAT)
 (sometimes called the "thunk table")</li>
 <li>
 <p align="left">Access CPU registers</li>
 <li>
 <p align="left">Increase your program's stack</li>
 <li>
 <p align="left">Change your program's entry point</li>
 <li>
 <p align="left">Increase the maximum number of modules</li>
 <li>
 <p align="left">Call procedures by address</li>
</ol>
<p align="left"><b>Required Tools</b></p>
<p align="left">In order to perform the presented objectives, you will need the
following:</p>
<ul>
 <li>
 <p align="left">Visual Basic 5.0 or 6.0 (sorry, VB.NET doesn't work with
 this code)</li>
 <li>
 <p align="left">A C compiler, preferably Visual C++</li>
 <li>
 <p align="left">A debugger, such as SoftIce (if you don't want to spend the
 money or time downloading a debugger, you'll be able to write your own after
 reading this article)</li>
 <li>
 <p align="left">An assembler, preferably Macro Assembler (MASM)</li>
</ul>
<hr>
<p align="left"><b>Background Information You Need To Read</b></p>
<p align="left">Despite what people may think, Visual Basic isn't a true
language.  What many people don't understand is that Visual Basic's
compiler only generates native code.  This gives your programs better
performance, and above all, bullet-proof security for your source.  After
all, how many VB5 and VB6 decompilers do <i>you </i>know of?  All this
means you have less control over how your binary programs are complied, which
can give you a major headache when you want to keep the number of dependent
files to a bare minimum.  Alas, all is not lost.  You now have the
power to seize control of Visual Basic and give it back to your program. 
As you read, you will be able to intercept VB's native code generation and link
custom object modules into your project</p>
<p align="left">However, this after-the-fact added availability has a
forewarning that is worth mentioning: Microsoft will NOT like the idea that
there are programs out there that can now intercept internal API calls of the VB
environment (and most of Windows for that matter).  This rules out giving
you access to compiler.  But that is exactly what this article and code
accomplishes.</p>
<blockquote>
 <p align="left"><font color="#FF0000"><b>**CRASH-YOUR-COMPUTER WARNING** </b>You
 can safely view the assembly source code of your projects using this add-in,
 but you can count on seeing a <i>lot</i> of General Protection Faults if you
 use the add-in to start inserting your own C or assembly code in a VB
 binary.  I'm not saying it shouldn't be done, but I am saying you need to
 consider the power vs. danger trade-off carefully, as you do with any advanced
 technique.</font></p>
</blockquote>
<p align="left"><b>Basic Info On The Visual Basic Compiler and How To Harness It</b></p>
<p align="left">VB's compiler consists of two programs: C2.exe and Link.exe. 
Link.exe does just that: it links your object code with intermediate library
code and writes the executable.  C2 is an older version of Microsoft's
second-pass C compiler; Microsoft modified it specifically for use with VB, and
it is called once for every file in your project.</p>
<p align="left">C2 and Link are activated with the kernel function CreateProcess. 
This is where the magic starts.  By hooking the CreateProcess API call, you
are able to intercept and modify commands sent to C2 and Link.  You're
probably thinking "How in the heck do you hook an API call in a VB
program?"  Indeed, this process is complex to say the least, but if
NuMega can do it with SoftIce, you can do it with Visual Basic.</p>
<p align="left"><b>Final Notes Before Downloading the Code</b></p>
<p align="left">I <b>strongly</b> recommend reading the original article by John
Chamberlain (which is included in the ZIP), following it step-by-step, and reading
it very carefully until you really understand what's going on. Once you understand how the controller works, you will find it easy to
use; if you skip ahead, you may experience frustration. It goes without saying that this is a sophisticated tool that is appropriate<i> only for advanced programmers.</i> When you use it, you leave the world of the help file behind and enter into uncharted territory. The challenges and risks of forging into this wilderness are substantial, but the potential reward is well worth it: nearly total control over your VB executable.</p>
<p align="left">Microsoft includes an assembler called ML.EXE in its Win98 DDK,
which is available for download at <a href="http://www.microsoft.com/ddk/ddk98.htm">http://www.microsoft.com/ddk/ddk98.htm</a>. Theoretically, you can buy MASM from Microsoft, but I could not find out how to buy it. You might have to have wax one of Bill's cars or something before they sell it to you. Microsoft seems to be adopting the same position toward assembly that the government has towards uranium.</p>
<p align="left">You won't get far with the Compile Controller unless you have a working knowledge of assemblers and assembly language. If the last program you assembled was on punched cards, now wouldn't be a
bad time to brush up on your skills. I found the printed (yes, printed!) MASM 6.1 manuals invaluable for this purpose. You will also absolutely need a programmer's reference manual on the x86 instruction set. To get this, call (800) 548-4725 (the Intel literature distribution center). The best book on x86 assembly in print that is easily available is Master Class Assembly Language, but this book is in no way a substitute for the MASM manuals. Check out the assembly language newsgroups and their FAQs for more information. Also, note that the Microsoft knowledge base has a number of useful articles on mixed language development that are relevant.</p>
<p align="center"><b>Now go forth and kick tail, programmer!</b></p>

