import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

FileName = int(input("What should the file be saved as?: "))

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(f"{FileName}.xlsx")
worksheet = workbook.add_worksheet()

PrevTotalFormat = workbook.add_format()
PrevTotalFormat.set_right(5)

RolledValue = workbook.add_format()
RolledValue.set_bottom(5)

PossibleResults = workbook.add_format()
PossibleResults.set_border(5)
PossibleResults.set_border_color("green")

PosibilityCount = workbook.add_format()
PosibilityCount.set_border(5)
PosibilityCount.set_left(0)
PosibilityCount.set_border_color("orange")

PosibilityCount = workbook.add_format()
PosibilityCount.set_border(5)
PosibilityCount.set_left(0)
PosibilityCount.set_border_color("orange")

PercentageFormat = workbook.add_format({'num_format': '0.00%'})

PercentageConditionalFormating ={'type':     '3_color_scale',
                      'min_type': 'min',
                      'mid_type': 'percent',
                      'max_type': 'max',
                      'min_value': '',
                      'mid_value': '50',
                      'max_value': '',
                      'min_color': '#63BE7B',
                      'mid_color': '#FFEB84',
                      'max_color': '#F8696B',}

DiceSize = int(input("How many sides should the rolled dice have?: "))

DiceInputs = int(input("How many dice should be rolled?: "))

for YCoordinates in range(0, (DiceSize)):
    worksheet.write(YCoordinates+1, 0, YCoordinates+1, PrevTotalFormat)
    worksheet.write(YCoordinates+1, 1, YCoordinates+1)

if DiceSize < 8:
    BlockSize = 8
else:
    BlockSize = DiceSize+1

MaxHeight = ((DiceSize-1)*(DiceInputs-1))+3
worksheet.write(0, 1, 0,RolledValue)


# Creates the data blocks from the first through to the point where the values it belives are rolled exceed the limit of the die.  
for Generation in range(1, 3): 
    TableSize = ((DiceSize-1)*Generation)+2
    for row in range(1, 2):
        for column in range((BlockSize*Generation)+1,(BlockSize*Generation)+DiceSize+1):
            worksheet.write(row-1, column, column-(BlockSize*Generation),RolledValue)
    for row in range(1, TableSize):
        # This creates the numbering of each row in the upper data tables.
        for column in range(1+(BlockSize*Generation),2+(BlockSize*Generation)):
            worksheet.write(row, column-1, row+Generation-1,PrevTotalFormat)
        
        # This creates the upper data decimal spread.
        for column in range(1+(BlockSize*Generation),(DiceSize+1)+(BlockSize*Generation)):
            worksheet.write(row, column, (row+Generation-1)+(column-(Generation*BlockSize)))
        
            
    for row in range(MaxHeight, MaxHeight+(TableSize-1)):
        for column in range(BlockSize*(Generation-1),1+(BlockSize*(Generation-1))):
            worksheet.write(row, column, (row-MaxHeight)+Generation,PossibleResults)
        for column in range(1+(BlockSize*(Generation-1)),2+(BlockSize*(Generation-1))):
            # This creates the lower data count of possible ways to reach a given sum
            for Count in range(1, DiceSize):
                worksheet.write(Count+MaxHeight-1, column, Count, PosibilityCount)
            for DownCount in range(1, DiceSize+1):
                worksheet.write((TableSize+MaxHeight)-(DownCount+1), column, DownCount, PosibilityCount)
            for TopCount in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
                worksheet.write(MaxHeight+TopCount-1, column, DiceSize, PosibilityCount)
            

        # This creates the representation of the lower data as a decimal chance of getting a given sum.  
        for column in range(2+(BlockSize*(Generation-1)),3+(BlockSize*(Generation-1))):
            Top = xl_rowcol_to_cell(MaxHeight,1+(BlockSize*(Generation-1)))
            Bottom = xl_rowcol_to_cell(MaxHeight+(TableSize-2),1+(BlockSize*(Generation-1)))                 
            worksheet.write_formula(row, column, f"={xl_rowcol_to_cell(row, column-1)}/SUM({Top}:{Bottom})")
  
        # This creates the representation of the lower data as a percentage chance of getting a sum >= the given number.
        for column in range(4+(BlockSize*(Generation-1)),5+(BlockSize*(Generation-1))):
            Top = xl_rowcol_to_cell(row,3+(BlockSize*(Generation-1)))
            Bottom = xl_rowcol_to_cell(MaxHeight+(TableSize-2),3+(BlockSize*(Generation-1)))               
            worksheet.write_formula(row, column, f"=SUM({Top}:{Bottom})")

    # This creates the representation of the lower data as a percentage chance of getting a given sum.
    for column in range(3+(BlockSize*(Generation-1)),4+(BlockSize*(Generation-1))):
        for AsPercent in range(1, DiceSize+1):
            worksheet.write_formula(AsPercent+MaxHeight-1, column, f"={xl_rowcol_to_cell(AsPercent+MaxHeight-1, column-1)}",PercentageFormat)
            
        for DownPercent in range(1, DiceSize):
            worksheet.write_formula((TableSize+MaxHeight)-(DownPercent+1), column, f"={xl_rowcol_to_cell((TableSize+MaxHeight)-(DownPercent+1), column-1)}",PercentageFormat)
            
        for TopPercent in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
            worksheet.write_formula(MaxHeight+TopPercent-1, column, f"={xl_rowcol_to_cell(MaxHeight+TopPercent-1, column-1)}",PercentageFormat)

        worksheet.conditional_format(MaxHeight, column, (TableSize+MaxHeight)+1, column, PercentageConditionalFormating)

# Creates upper data blocks from generation 3 through to the second last.
for Generation in range(3, DiceInputs): 
    TableSize = ((DiceSize-1)*Generation)+2
    for row in range(1, 2):
        for column in range((BlockSize*Generation)+1,(BlockSize*Generation)+DiceSize+1):
            worksheet.write(row-1, column, column-(BlockSize*Generation),RolledValue)
    for row in range(1, TableSize):
        for column in range(1+(BlockSize*Generation),2+(BlockSize*Generation)):
            worksheet.write(row, column-1, row+Generation-1,PrevTotalFormat)
    for row in range(1, (TableSize-DiceSize+1)):
        for column in range((BlockSize*(Generation-1))+1,(BlockSize*(Generation-1))+DiceSize+1):
            worksheet.write_formula(row, column, f"={xl_rowcol_to_cell((row+MaxHeight-1), ((Generation-2)*BlockSize)+2)}*(1/6)")
            
    for row in range(MaxHeight, MaxHeight+(TableSize-1)):
        for column in range(BlockSize*(Generation-1),1+(BlockSize*(Generation-1))):
            worksheet.write(row, column, (row-MaxHeight)+Generation,PossibleResults)
        for column in range(2+(BlockSize*(Generation-1)),3+(BlockSize*(Generation-1))):
            Top = xl_rowcol_to_cell(MaxHeight,1+(BlockSize*(Generation-1)))
            Bottom = xl_rowcol_to_cell(MaxHeight+(TableSize-2),1+(BlockSize*(Generation-1)))
            for ProbCount in range(0, DiceSize-1):
                ProbCountFormula = "0"
                for FormulaIterable in range(0, ProbCount+1):
                    ProbCountFormula += f"+{xl_rowcol_to_cell(ProbCount-FormulaIterable+1, 1+FormulaIterable+(BlockSize*(Generation-1)))}"
                worksheet.write_formula(ProbCount+MaxHeight, column, f"={ProbCountFormula}")
            for ProbCount in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
                ProbCountFormula = "0"
                for FormulaIterable in range(0, DiceSize):
                    ProbCountFormula += f"+{xl_rowcol_to_cell(ProbCount-FormulaIterable, 1+FormulaIterable+(BlockSize*(Generation-1)))}"
                worksheet.write_formula(ProbCount+MaxHeight-1, column, f"={ProbCountFormula}")

            for ProbCount in range(1, DiceSize):
                ProbCountFormula = "0"
                for FormulaIterable in range(0, (DiceSize-ProbCount)):
                    ProbCountFormula += f"+{xl_rowcol_to_cell((((DiceSize-1)*(Generation-1))+2)-(FormulaIterable+1), 1+ProbCount+FormulaIterable+(BlockSize*(Generation-1)))}"
                worksheet.write_formula(MaxHeight+TableSize-(DiceSize-ProbCount)-1, column, f"={ProbCountFormula}")
        # This creates the representation of the lower data as a percentage chance of getting a given sum.
        for column in range(3+(BlockSize*(Generation-1)),4+(BlockSize*(Generation-1))):
            for AsPercent in range(1, DiceSize):
                worksheet.write_formula(AsPercent+MaxHeight-1, column, f"={xl_rowcol_to_cell(AsPercent+MaxHeight-1, column-1)}", PercentageFormat)
                
            for DownPercent in range(1, DiceSize):
                worksheet.write_formula((TableSize+MaxHeight)-(DownPercent+1), column, f"={xl_rowcol_to_cell((TableSize+MaxHeight)-(DownPercent+1), column-1)}",PercentageFormat)
                
            for TopPercent in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
                worksheet.write_formula(MaxHeight+TopPercent-1, column, f"={xl_rowcol_to_cell(MaxHeight+TopPercent-1, column-1)}", PercentageFormat)

            worksheet.conditional_format(MaxHeight, column, (TableSize+MaxHeight)+1, column, PercentageConditionalFormating)

        # This creates the representation of the lower data as a percentage chance of getting a sum >= the given number.
        for column in range(4+(BlockSize*(Generation-1)),5+(BlockSize*(Generation-1))):
            Top = xl_rowcol_to_cell(row,3+(BlockSize*(Generation-1)))
            Bottom = xl_rowcol_to_cell(MaxHeight+(TableSize-2),3+(BlockSize*(Generation-1)))               
            worksheet.write_formula(row, column, f"=SUM({Top}:{Bottom})")
            
    # This creates the lower data count of possible ways to reach a given sum       
    for column in range(1+(BlockSize*(Generation-1)),2+(BlockSize*(Generation-1))):
        for Count in range(1, DiceSize+1):
            worksheet.write(Count+MaxHeight-1, column, Count,PosibilityCount)
        for DownCount in range(1, DiceSize+1):
            worksheet.write((TableSize+MaxHeight)-(DownCount+1), column, DownCount,PosibilityCount)
        for TopCount in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
            worksheet.write(MaxHeight+TopCount-1, column, DiceSize,PosibilityCount)
    

    for column in range(1,2):
        for Count in range(1, DiceSize):
            worksheet.write(Count+MaxHeight-1, column, 1,PosibilityCount)



Generation = DiceInputs
TableSize = ((DiceSize-1)*Generation)+2
# Creates data blocks for the last generation.
for row in range(MaxHeight, MaxHeight+(TableSize-1)):

    for column in range(BlockSize*(Generation-1),1+(BlockSize*(Generation-1))):
        worksheet.write(row, column, (row-MaxHeight)+Generation,PossibleResults)

    # This creates the representation of the lower data as a percentage chance of getting a sum >= the given number.
    for column in range(4+(BlockSize*(Generation-1)),5+(BlockSize*(Generation-1))):
        Top = xl_rowcol_to_cell(row,3+(BlockSize*(Generation-1)))
        Bottom = xl_rowcol_to_cell(MaxHeight+(TableSize-2),3+(BlockSize*(Generation-1)))           
        worksheet.write_formula(row, column, f"=SUM({Top}:{Bottom})")
        
    for column in range(2+(BlockSize*(Generation-1)),3+(BlockSize*(Generation-1))):
        # This creates the representation of the lower data as a decimal chance of getting a given sum.  
        Top = xl_rowcol_to_cell(MaxHeight,1+(BlockSize*(Generation-1)))
        Bottom = xl_rowcol_to_cell(MaxHeight+(TableSize-2),1+(BlockSize*(Generation-1)))
        for ProbCount in range(0, DiceSize-1):
            ProbCountFormula = "0"
            for FormulaIterable in range(0, ProbCount+1):
                ProbCountFormula += f"+{xl_rowcol_to_cell(ProbCount-FormulaIterable+1, 1+FormulaIterable+(BlockSize*(Generation-1)))}"
            worksheet.write_formula(ProbCount+MaxHeight, column, f"={ProbCountFormula}")
        for ProbCount in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
            ProbCountFormula = "0"
            for FormulaIterable in range(0, DiceSize):
                ProbCountFormula += f"+{xl_rowcol_to_cell(ProbCount-FormulaIterable, 1+FormulaIterable+(BlockSize*(Generation-1)))}"
            worksheet.write_formula(ProbCount+MaxHeight-1, column, f"={ProbCountFormula}")
        for ProbCount in range(1, DiceSize):
            ProbCountFormula = "0"
            for FormulaIterable in range(0, (DiceSize-ProbCount)):
                ProbCountFormula += f"+{xl_rowcol_to_cell((((DiceSize-1)*(Generation-1))+2)-(FormulaIterable+1), 1+ProbCount+FormulaIterable+(BlockSize*(Generation-1)))}"
            worksheet.write_formula(MaxHeight+TableSize-(DiceSize-ProbCount)-1, column, f"={ProbCountFormula}")    

    # This creates the lower data count of possible ways to reach a given sum
    for column in range(1+(BlockSize*(Generation-1)),2+(BlockSize*(Generation-1))):
        for Count in range(1, DiceSize):
            worksheet.write(Count+MaxHeight-1, column, Count,PosibilityCount)
        for DownCount in range(1, DiceSize):
            worksheet.write((TableSize+MaxHeight)-(DownCount+1), column, DownCount,PosibilityCount)
        for TopCount in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
            worksheet.write(MaxHeight+TopCount-1, column, DiceSize,PosibilityCount)

    # This calculates the chance of getting a specific roll on top of a given total.
    for row in range(1, TableSize-(DiceSize-1)):
        for column in range(1+(BlockSize*(Generation-1)),DiceSize+(1+(BlockSize*(Generation-1)))):
            worksheet.write_formula(row, column, f"={xl_rowcol_to_cell((row+MaxHeight-1), ((Generation-2)*BlockSize)+2)}*(1/6)")       

    # This creates the representation of the lower data as a percentage chance of getting a given sum.
    for column in range(3+(BlockSize*(Generation-1)),4+(BlockSize*(Generation-1))):
        for AsPercent in range(1, DiceSize+1):
            worksheet.write_formula(AsPercent+MaxHeight-1, column, f"={xl_rowcol_to_cell(AsPercent+MaxHeight-1, column-1)}", PercentageFormat)
            
        for DownPercent in range(1, DiceSize+1):
            worksheet.write_formula((TableSize+MaxHeight)-(DownPercent+1), column, f"={xl_rowcol_to_cell((TableSize+MaxHeight)-(DownPercent+1), column-1)}", PercentageFormat)
            
        for TopPercent in range(DiceSize,((DiceSize-1)*(Generation-1))+2):
            worksheet.write_formula(MaxHeight+TopPercent-1, column, f"={xl_rowcol_to_cell(MaxHeight+TopPercent-1, column-1)}", PercentageFormat)

        worksheet.conditional_format(MaxHeight, column, (TableSize+MaxHeight)+1, column, PercentageConditionalFormating)


workbook.close()
