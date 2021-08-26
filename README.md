# random-id-generator
## Intro
Some library code to generate random IDs to assign to database records. Originally written for Access database. Collision-detection is necessary (so either a table constraint or similar). The user can control the characters used, ID length, etc.

## Parameters
Parameters include number of characters in the generated IDs (length) and character set used. The default character set contains 31 uppercase ASCII characters (numbers, letters), excluding look-alikes (ex., "l" (lowercase El) and "I" (uppercase Eye), "0" (zero) and "O" (uppercase Oh). This makes the IDs more suitable for data entry tasks.

ID length determines the number of unique IDs able to be generated, and also the number of duplicates generated. If the number of IDs to generate is a large fraction of the ID space, more duplicates will result in increased runtime. In this case, increasing the ID length by one will greatly reduce this overhead. When the code is run, it prints details (including runtime and number of duplicates that needed to be regenerated) to the console.

If your use case makes it unacceptable to generate and attempt to save duplicate IDs into your target table, consider making a local working table for this task.

Example: An ID length of (only) 4 characters, using the default character set (31 chars), allows for 923,521 permutations. (To explore more combinations of parameters, try the [mathisfun.com calculator](https://www.mathsisfun.com/combinatorics/combinations-permutations-calculator.html).)