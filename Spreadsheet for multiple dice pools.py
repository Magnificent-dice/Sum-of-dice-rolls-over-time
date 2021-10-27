import math
from math import log10, floor

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

Attempt = input("Attempt number: ")

workbook = xlsxwriter.Workbook(f"{Attempt}.xlsx")
worksheet = workbook.add_worksheet()

DamageFormat = workbook.add_format()
DamageFormat.set_bottom(5)

TurnFormat = workbook.add_format()
TurnFormat.set_right(5)

CloseToFifty = workbook.add_format()
CloseToFifty.set_font_color('#9C0006')
CloseToFifty.set_bg_color('#FFC7CE')

ChanceList = []
TurnList = []

DicePool = int(input("What is the weapon's batch size: "))
DiceSize = int(input("What is the weapon's damage: "))
DicePerShot = int(input("How many dice does a shot deal: "))

ListLimit = 10


def sum_chance(TargetValue, DicePool, DiceSize):
    Summation = 0
    Iteration = 0
    for KValues in range(0,math.floor((TargetValue-DicePool)/DiceSize)+1):
        AlternatingSign = (-1)**KValues
        NChooseK = math.comb(DicePool, KValues)

        NLargeBinomeal = (TargetValue-DiceSize*KValues-1)
        KLargeBinomeal = (TargetValue-DiceSize*KValues-DicePool)
        LargerBinomeal = math.comb(NLargeBinomeal, KLargeBinomeal)
        Summation += (AlternatingSign*NChooseK*LargerBinomeal)

    return Summation


def entry_for_list(RelevantList, DicePool, DiceSize, DicePerShot):
    for Values in range(1,(DicePool*DiceSize*DicePerShot)+1):
        RelevantList.append(sum_chance(Values, DicePool*DicePerShot, DiceSize))
    return RelevantList


PoolSizeByTurn = []

for Turns in range(1,ListLimit+1):
    PoolSizeByTurn.append(entry_for_list([], Turns, DicePool, DicePerShot))




PoolSizeChanceByTurn = []

for Turns in range(1,ListLimit+1):
    TurnChances = []
    for Damage in PoolSizeByTurn[Turns-1]:
        ChanceOfPool = (Damage*(1/(sum(PoolSizeByTurn[Turns-1]))))
        TurnChances.append(ChanceOfPool)
    PoolSizeChanceByTurn.append(TurnChances)


for Turns in range(1,ListLimit+1):
    ResultsApplied = []
    for PossiblePools in range(1,len(PoolSizeChanceByTurn[Turns-1])+1):
        PoolList = entry_for_list([], PossiblePools, DiceSize, DicePerShot)
        ResultsApplied.append(PoolList)
    TurnList.append(ResultsApplied)
    
DamageChanceByTurn = []
for Turns in range(0,ListLimit):
    GivenTurnDamage = []
    DamageChanceByTurn.append(GivenTurnDamage)
    for ResultsApplied in range(0,len(TurnList[Turns])):
        ResultsSum = sum(TurnList[Turns][ResultsApplied])
        for Damage in range(0,len(TurnList[Turns][ResultsApplied])):
            DamageChance = (TurnList[Turns][ResultsApplied][Damage]/ResultsSum)*PoolSizeChanceByTurn[Turns][ResultsApplied]
            if len(DamageChanceByTurn[Turns]) <= Damage:
                DamageChanceByTurn[Turns].append(DamageChance*100)
            else:
                DamageChanceByTurn[Turns][Damage] += (DamageChance*100)
            
for Turns in range(0,ListLimit):
    for ResultsApplied in range(0,len(TurnList[Turns])):
        for Damage in range(0,len(TurnList[Turns][ResultsApplied])):
            DamageChanceByTurn[Turns][Damage] = round(DamageChanceByTurn[Turns][Damage],3)
print()

Counter = 1
for Turns in DamageChanceByTurn:
    print()
    print(f"Turns[{Counter}]: {Turns}")
    Counter += 1

YCounter = 1
for Turns in DamageChanceByTurn:
    XCounter = 1
    for Damage in Turns:
        worksheet.write(YCounter, XCounter, Damage)
        XCounter += 1
    worksheet.conditional_format(YCounter, 1, YCounter, len(Turns),{'type':     '3_color_scale',
                                                                     'min_type': 'min',
                                                                     'mid_type': 'percent',
                                                                     'max_type': 'max',
                                                                     'min_value': '',
                                                                     'mid_value': '50',
                                                                     'max_value': '',
                                                                     'min_color': '#63BE7B',
                                                                     'mid_color': '#FFEB84',
                                                                     'max_color': '#F8696B',})
                                                                    
    YCounter += 1

YCounter = 1
for Turns in DamageChanceByTurn:
    XCounter = 1
    for Damage in Turns:
        Start = xl_rowcol_to_cell(YCounter, XCounter)
        End = xl_rowcol_to_cell(YCounter, len(Turns))
        worksheet.write(YCounter+len(DamageChanceByTurn)+1, XCounter, f"=SUM({Start}:{End})")
        XCounter += 1
    YCounter += 1
    
worksheet.conditional_format(1+len(DamageChanceByTurn), 1, 1+(len(DamageChanceByTurn)*2), len(Turns),{'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  45,
                                       'maximum':  55,
                                       'format':   CloseToFifty})
    
for Damage in range(1,len(DamageChanceByTurn[-1])+1):
    worksheet.write(0, Damage, Damage, DamageFormat)
for Turns in range(1,len(DamageChanceByTurn)+1):
    worksheet.write(Turns, 0, Turns, TurnFormat)

workbook.close()
