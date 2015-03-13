## Steps to integrate SpellCode into the Visual Studio IDE ##

Please follow the below mentioned steps to add SpellCode to the Microsoft Visual Studio IDE as an external tools.

1. Open the Microsoft Visual Studio IDE 

&lt;BR&gt;


2. Click _Tools_ > _External Tools..._ menu item 

&lt;BR&gt;


3. Press "_Add_" button 

&lt;BR&gt;


4. Specify the following parameters, 

&lt;BR&gt;



  * **Title** : SpellCode
  * **Command** : specify the full filename and the path to the spellcode.exe file
  * **Arguments** : "$(ProjectDir)" 1
> > (In here 1 indicates the C++ programming language, you can specify any value from 1 to 11, and it depends on your default programming language) 

&lt;BR&gt;


5. Press "_OK_" button.