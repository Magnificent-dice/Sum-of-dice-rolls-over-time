# Sum-of-dice-rolls-over-time
This repository covers the python code I made to automatically model the results of a series of dice rolls. Originally designed for my homebrewed d&d naval battle system this code generates the chances of getting specific sums from rolling pools of dice. It comes in multiple forms for multiple kinds of results modelling.

The first program represents a pool of dice as each die is rolled and added to the total. It creates a .xlsx file with each dice roll represented as a block with previous totals along the vertical, and the new result along the top. It also represents the data with a table of each total, the number of results that produce it, its probability as a decimal, as a percentage, and the chance of getting at least that number.

The second program represents a pool of dice which is of a size determined by another die roll. It creates a .xlsx file with the horizontal being the sum of all dice results, and the verticle being the amount of times the entire lot has been rolled. The upper half has the chance that specific sum from that many rolls, and the lower half has the chance that a result equal to or greater than that sum will be rolled.
