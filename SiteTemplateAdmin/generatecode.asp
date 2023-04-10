<%
 FUNCTION GetRandomCode(codelength,numberofcombinations)
dim codeCharacters, codeArray,x,thiscode,totalcode
'Number of characters in the array below
codecharacters = 35

'Array of characters being used for the random code
codearray = Array("a","b","c","d","e","f","g","h","i","j","k","l", _
"m","n","o","p","q","r","s","t","u","v","w","x", _
"y","z","1","2","3","4","5","6","7","8","9")

'Generates one random character until it reaches code length
FOR x = 1 TO codelength
RANDOMIZE
'Gets a random number based on the value in the codecharacters variable
thiscode = (Int(((codecharacters - 1) * Rnd) + 1))

'builds the code on top of itself until complete by selecting the
'character from the array based on the random number generated above
totalcode = totalcode & codearray(thiscode)

'This is only an extra thing to tell you how many combinations
'there are for the code length specified. You may get scientific
'notation to describe the length so, just take it that is damn near
'impossible to guess the code.
IF numberofcombinations = "" THEN numberofcombinations = 1
numberofcombinations = numberofcombinations * codecharacters

NEXT

'Random code that is returned after codelength has been fulfilled
Getrandomcode = totalcode

END FUNCTION
%>