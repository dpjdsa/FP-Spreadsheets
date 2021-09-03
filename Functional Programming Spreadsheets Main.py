# Main Code to decode AST including returns list of code and formula. Decodes list(range(a,b,c)) construct
import pdb
import ast
import re
import random
from params import *
from treelib import Node,Tree
from datetime import datetime

#
# Reads in function definitions from file filename and returns list of function definitions
def readfunctions(filename):
    funclist=[]
    with open(filename) as f:
        while True:
            inputtxt=f.readline()
            if not inputtxt:
                break
            # Skip any comment lines at beginning which start with #
            if inputtxt[0]!='#':
                funclist.append(inputtxt)
    print("Function Lines Read in are: ",funclist)
    f.close()
    return funclist
# Looks in formula for cells of type A-Z,0-9 in formula string and shifts them down by inc, acts recursively
def shift_formula_down(formula,inc):
    # Check for cell references of type [A-Z][0-9] which therefore have no absolute row references
    p=re.compile('[A-Z][0-9]')
    m=p.search(formula)
    # If any such reference found then increment the row part of the reference by inc and analyse remainder recursively
    if m:
        return formula[:m.start()+1]+str(int(formula[m.start()+1])+inc)+shift_formula_down(formula[m.start()+2:],inc)
    else:
        return formula
# Checks for all types of keywords which are allowable for decoding
def permitted_parameter(value):
    return value in ('name','body','args','arg','left','right','op','value','id','n','func')
# Code to output results as CSV file which can be read into Excel
def opsheetCSV(name,dict,formulas,retexp):
    #Open file
    print("Writing Output CSV to testfile.csv")
    f=open("testfile.csv","w")
    # Write row including title,date,time and function name in first row
    row1="Test Functional Programme to Spreadsheet"
    # Get timestamp and date and write it and number of folds to sheet
    now=datetime.now()
    row1=row1+","+now.strftime("%d/%m/%Y %H:%M:%S")+","+"FOLDS: "+str(NUMFOLDS)+","
    rightcol=3
    # Allow spare columns for arguments before writing function call out to sheet
    for i in range(1,len(dict)):
        row1+=","
        rightcol+=1
    # Now create a reverse dictionary to ensure arguments are inserted into correct cell positions
    revdict={}
    # Create reverse dictionary for absolute references only, stripping out all $ characters from the references
    # This is so only the arguments which appear in the parameter list are output to the CSV file
    for key,value in dict.items():
        if value[0][0]=="$":
            revdict[value[0].replace("$","")]=key
    print("Reverse Dictionary=",revdict)
    # Get a sorted list of keys of the reverse dictionary
    sortedkeys=sorted(list(revdict.keys()))
    # Now assemble arguments to function adding commas between arguments and removing last comma before close bracket
    name+="("
    for i in sortedkeys:
        name+=revdict[i]+","
    name=name[:-1]
    name+=")"
    row1=row1+'"'+"FUNCTION NAME: "+name+'"'+","+'"'+"RETURNS: "+retexp+'"'
    rightcol+=2
    for i in range(rightcol,ord(MAXCOL)-64):
        row1+=","
    row1+=chr(124)+"\n"
    print("Row 1:=",row1)
    f.write(row1)
    # Write variables in second row
    row2="Variables:"
    rightcol=1
    # Put the variable names and values into the columns corresponding to references in the sorted keys
    curcol=1
    row3="Values:"
    for key in sortedkeys:
        while True:
            # Check if column name matches column of argument cell reference
            if chr(65+curcol)==key[0]:
                break
            # If not skip to next column in both rows 2 and 3
            row2+=","
            row3+=","
            curcol+=1
            rightcol+=1
        # Column matches cell reference so put argument name and value in column of rows 2 and 3.
        row2+=","+revdict[key]
        row3+=","+str(Argdict[revdict[key]][1])
        curcol+=1
        rightcol+=1
    # Add column descriptions after blank column and record start of formulas in startformcol
    startformcol=curcol
    row2+=","
    rightcol+=1
    for description in Desccol:
        row2=row2+","+'"'+description+'"'
        rightcol+=1
    # and add description of return value
    for i in range(rightcol,ord(MAXCOL)-64):
        row2+=","
    row2=row2+chr(124)+'\n'
    print("Row 2=:",row2)
    f.write(row2)
    # Add blank column before writing out formulas
    row3+=","
    for item in formulas:
        row3+=","+'"'+item+'"'
    for i in range(rightcol,ord(MAXCOL)-64):
        row3+=","
    row3+=chr(124)+'\n'
    print("Row 3=:",row3)
    f.write(row3)
    # Copy formulas down according to settings in Copydown list
    for i in range(2,NUMFOLDS):
        row=""
        for j in range(startformcol):
            row+=','
        rightcol=startformcol+1
        for (offset,item) in enumerate(formulas):
            if Copydown[offset]:
                row+=","+'"'+shift_formula_down(item,(i-1))+'"'
            else:
                row+=","
            rightcol+=1
        for j in range(rightcol,ord(MAXCOL)-64):
            row+=","
        row+=chr(124)+'\n'
        print("Row "+str(i+2)+"=:",row)
        f.write(row)
    endrow=""
    # For last row add formula which will fill each cell width with ~ characters
    for i in range(ord(MAXCOL)-65):    
        endrow+='"=REPT(CHAR(126),CELL(""WIDTH"",'+chr(65+i)+'1))",'
    endrow+=chr(42)
    print("Row "+str(NUMFOLDS+2)+"=:",endrow)
    f.write(endrow)
    f.close()
    return

# Lists the node and unpacks the children of a node
def str_node(node):
    if isinstance(node, ast.AST):
        fields = [(name, str_node(val)) for name, val in ast.iter_fields(node) if permitted_parameter(name)]
        rv = '%s(%s' % (node.__class__.__name__, ', '.join('%s=%s' % field for field in fields))
        return rv + ')'
    else:
        return repr(node)
# Translates the Python entity represented by node into an Excel formula and description of code translated 
def Decode_Gen(node,argnumflg):
    # Argument Dictionary is global
    global Funcname,Argcol,Absflg,Writeflg
    # Check for Module
    if isinstance(node,ast.Module):
        print("Module Found")
        return "",[],[]
    # Check for Function Definition
    elif isinstance(node,ast.FunctionDef):
        d=dict(ast.iter_fields(node))
        Funcname=d['name']
        print("\t*** in Function Definition",Funcname,Funcname)
        return Funcname,Funcname,Funcname
    # Check for Arguments
    elif isinstance(node,ast.arguments):
        print("Arguments Found")
        # Get arguments
        d=dict(ast.iter_fields(node))
        # Create dictionary of Argument, Cell Reference, Random value (1-50)
        count=len(d['args'])
        opstring=""
        for i in range(len(d['args'])):
            d1=dict(ast.iter_fields(d['args'][i]))
            print("d1=",d1)
            opstring=opstring+d1['arg']
            if count>1:
                opstring+=","
            if d1['arg'] not in Argdict:
                if Absflg:
                    Argdict[d1['arg']]=("$"+Argcol+"$"+str(ARGROW),random.randrange(50)+1)
                else:
                    Argdict[d1['arg']]=(Argcol+str(ARGROW),random.randrange(50)+1)
                print("Key added:",d1['arg'],Argdict,Argcol)
                Argcol=chr(ord(Argcol)+1)
            else:
                print("Key ",d1['arg'],"already in Arg Dictionary")
            print("\t\t***In arguments, ArgCol=",Argcol)
        print("\t*** in Arguments",opstring,opstring)
        return opstring,opstring,opstring
    # Check for Lambda expression
    elif isinstance(node,ast.Lambda):
        print("Lambda Expression Found")
        d=dict(ast.iter_fields(node))
        print("Lambda Dictionary =",d)
        # Get argument from 1 column to the right
        Argcol=chr(ord(Argcol)+1)
        print("\t***In Lambda, ArgCol=",Argcol)
        print("Args decoded =",Decode_Gen(d['args'],False)[0])
        opstring,descstring,_=Decode_Gen(d['body'],False)
        # Add fix so that if formula returns error in the sheet then "" is returned
        if Writeflg:
            Formula.insert(0,"=IFERROR("+opstring+","+'"""")')
            Copydown.insert(0,True)
            descstring="Lambda "+Decode_Gen(d['args'],False)[1]+":("+descstring+")"
            Desccol.append(descstring)
            print("\t\t*** Formula Appended in Lambda, Formula is now",Formula)
        print("\t*** in Lambda",Formula[-1],descstring)
        return Formula[-1],descstring,descstring
    # Check for Return statement and derive Excel formula to calculate return value
    elif isinstance(node,ast.Return):
        print("Return Found")
        # Set a flag to ouput formulas and titles when processing the Return class
        Writeflg=True
        # Set all arguments found to have relative cell references (such as in Lambda functions)
        Absflg=False
        # Create Dictionary from the  Return object
        d=dict(ast.iter_fields(node))
        # and decode it
        opstring,descstring,_=Decode_Gen(d['value'],False)
        print("opstring=",opstring)
        # If the formula string is empty then Return clause contains the only formula to output
        if (not Formula):
            Formula.append("="+opstring)
            Copydown.append(False)
            Desccol.append(descstring)
        print("\t*** in Return",opstring,descstring)
        Writeflg=False
        return opstring,descstring,descstring
    elif isinstance(node,ast.Call):
        print("Call Found")
        d=dict(ast.iter_fields(node))
        print("Call Dictionary =",d)
        print("Decoded Func",d['func'])
        functype=Decode_Gen(d['func'],False)[0]
        print("Decoded General: ",functype)
        opstring=functype+"("
        if (functype=='range'):
            print("Need to address range here")
            descstring="Range("
            arglist=[]
            count=len(d['args'])
            start='0'
            step='1'
            stop='0'
            if count==1:
                stop,stopd,_=Decode_Gen(d['args'][0],False)
                descstring+=stopd+")"
            elif count==2:
                start,startd,_=Decode_Gen(d['args'][0],False)
                stop,stopd,_=Decode_Gen(d['args'][1],False)
                descstring+=startd+","+stopd+")"
            elif count==3:
                start,startd,_=Decode_Gen(d['args'][0],False)
                stop,stopd,_=Decode_Gen(d['args'][1],False)
                step,stepd,_=Decode_Gen(d['args'][2],False)
                descstring+=startd+","+stopd+","+stepd+")"
            print("Start=",start,"Stop=",stop,"Step=",step)
            for value in d['args']:
                nexttoken=Decode_Gen(value,False)[0]
                if (isinstance(nexttoken,int)):
                    opstring+=str(nexttoken)
                else:
                    opstring+=nexttoken
                arglist.append(nexttoken)
                if count>1:
                    opstring+=","
                else:
                    opstring+=")"
                count-=1
                print("Range Arglist=",arglist)
            rangeobj=RangeClass("$"+Argcol+"$"+str(ARGROW),start,stop,step)
            print("Range Object =",rangeobj)
            print("Range Evaluation, Start:",start," Stop:",stop, "Step:",step)
            print("\t*** in Range",opstring,descstring,rangeobj)
            return opstring,descstring,rangeobj
        elif (functype=='list'):
            print("Addressing list now")
            arglist=[]
            count=len(d['args'])
            print("In List:Number of args",count,"Args",d['args'])
            descstring=""
            for value in d['args']:
                print("\t*** In List, value =",value)
                nexttoken,descstring,nextobject=Decode_Gen(value,False)
                if isinstance(nextobject,RangeClass):
                    opstring=nextobject.makelist()
                    if Writeflg:
                        Formula.insert(0,"="+opstring)
                        Copydown.insert(0,True)
                        descstring="List("+descstring+")"
                        Desccol.insert(0,descstring)
                    print('\t\t*** Formula appended in List Range, Formula list is:',Formula)
                elif isinstance(nextobject,FilterClass):
                    opstring,opstring1=nextobject.makelist()
                    if Writeflg:
                        Formula.append("="+opstring)
                        Formula.append("="+opstring1)
                        Copydown.append(True)
                        Copydown.append(True)
                        descstring="List("+descstring+")"
                        Desccol.append("__ref__")
                        Desccol.append(descstring)
                    print('\t\t*** Formula appended in List Filter, Formula list is:',Formula)
                elif isinstance(nextobject,MapClass):
                    opstring=nextobject.makelist()
                    if Writeflg:
                        Formula.append("="+opstring)
                        Copydown.append(True)
                        descstring="List("+descstring+")"
                        Desccol.append(descstring)
                    print('\t\t*** Formula appended in List Map, Formula list is:',Formula)
            print("\t*** in list",Formula[-1],descstring)
            return Formula[-1],descstring,descstring
        elif (functype=='filter'):
            print("Addressing filter now")
            descstring="Filter("
            arglist=[]
            print("In Filter: Number of args",len(d['args']),"Args",d['args'])
            for value in d['args']:
                nexttoken,descstring1,_=Decode_Gen(value,False)
                if (isinstance(nexttoken,int)):
                    opstring+=str(nexttoken)+","
                else:
                    opstring+=nexttoken+","
                descstring+=descstring1+","
                arglist.append(nexttoken)
                print("Filter Arglist =",arglist)
            opstring=opstring[:-1]+")"
            descstring=descstring[:-1]+")"
            filterobj=FilterClass(chr(ord(Argcol)+1)+str(ARGROW))
            print ("Filter Object =",filterobj)
            if Writeflg:
                Formula.append("=IFERROR(IF("+Argcol+str(ARGROW)+","+chr(ord(Argcol)-1)+str(ARGROW)+',""""),"""")')
                Copydown.append(True)
                Desccol.append(descstring)
            print("\t*** in Filter",opstring,filterobj)
            return opstring,descstring,filterobj
        elif (functype=='map'):
            print("Addressing map now")
            descstring="Map("
            arglist=[]
            print("In Map: Number of args",len(d['args']),"Args",d['args'])
            for value in d['args']:
                nexttoken,descstring1,_=Decode_Gen(value,False)
                if (isinstance(nexttoken,int)):
                    opstring+=str(nexttoken)+","
                else:
                    opstring+=nexttoken+","
                descstring+=descstring1+","
                arglist.append(nexttoken)
                print("Map Arglist =",arglist)
            opsstring=opstring[:-1]+")"
            descstring=descstring[:-1]+")"
            mapobj=MapClass(chr(ord(Argcol)+1)+str(ARGROW))
            print ("Map Object =",mapobj)
            if Writeflg:
                Formula.append("="+Argcol+str(ARGROW))
                Copydown.append(True)
                Desccol.append(descstring)
            print("\t*** in Map",opstring,mapobj)
            return opstring,descstring,mapobj
        opstring=functype+'('
        count=len(d['args'])
        for value in d['args']:
            if count>1:
                opstring=opstring+str(Decode_Gen(value,True)[0])+','
            else:
                opstring=opstring+str(Decode_Gen(value,True)[0])+')'
            count-=1
        print("Call opstring:",opstring)
        return opstring,[],[]
    # Check for Name and decode it as a name or its value, Python variable name or Excel cell ref
    elif isinstance(node,ast.Name):
        return Decode_Name(node,argnumflg,False),Decode_Name(node,argnumflg,True),Decode_Name(node,argnumflg,True)
    # Check for Num and decode it
    elif isinstance(node,ast.Num):
        opstring=Decode_Num(node,argnumflg)
        return opstring,opstring,opstring
    # Check for common arithmetic and comparison operators except % which is decoded differently
    elif isinstance(node,ast.Add):
        return "+","+","+"
    elif isinstance(node,ast.Sub):
        return "-","-","-"
    elif isinstance(node,ast.Mult):
        return "*","*","*"
    elif isinstance(node,ast.Div):
        return "/","/","/"
    elif isinstance(node,ast.Pow):
        return "^","**","**"
    elif isinstance(node,ast.Eq):
        return "=","==","=="
    elif isinstance(node,ast.NotEq):
        return "<>","!=","!="
    elif isinstance(node,ast.Lt):
        return "<","<","<"
    elif isinstance(node,ast.LtE):
        return "<=","<=","<="
    elif isinstance(node,ast.Gt):
        return ">",">",">"
    elif isinstance(node,ast.GtE):
        return ">=",">=",">="
    elif isinstance(node,ast.USub):
        return"-","-","-"
    elif isinstance(node,ast.UAdd):
        return"+","+","+"
    # Check for comparison (Note this code only allows single level comparisons at the moment)
    elif isinstance(node,ast.Compare):
        print("Comparison Found")
        d=dict(ast.iter_fields(node))
        print("Comparison Dict =",d)
        opstringl,descstringl,_=Decode_Gen(d['left'],False)
        opstringo,descstringo,_=Decode_Gen(d['ops'][0],False)
        opstringc,descstringc,_=Decode_Gen(d['comparators'][0],False)
        opstring=opstringl+opstringo+opstringc
        descstring=descstringl+descstringo+descstringc
        return opstring,descstring,descstring
    # Check for UnaryOp and decode
    elif isinstance(node,ast.UnaryOp):
        print("UnaryOp Found")
        d=dict(ast.iter_fields(node))
        print("UnaryOp Dict= ",d)
        opstringo,descstringo,_=Decode_Gen(d['op'],False)
        opstringd,descstringd,_=Decode_Gen(d['operand'],False)
        opstring="("+opstringo+opstringd+")"
        descstring="("+descstringo+descstringd+")"
        print("\t*** in unaryop",opstring,descstring)
        return opstring,descstring,descstring
    # Check for BinOp and decode
    elif isinstance(node,ast.BinOp):
        #pdb.set_trace()
        print("BinOp Found")
        d=dict(ast.iter_fields(node))
        print("Binop Dict= ",d)
        opstringl,descstringl,_=Decode_Gen(d['left'],False)
        opstringr,descstringr,_=Decode_Gen(d['right'],False)
        #If function is % (mod operator) create Excel function MOD(Left,Right)
        if isinstance(d['op'],ast.Mod):
            print("Mod Found")
            print(d['left'],d['right'])
            opstring="MOD("+opstringl+","+opstringr+")"
            descstring="("+descstringl+"%"+descstringr+")"
        # Otherwise for all other operators decode as Left, Op, Right sequence
        else:
            opstring,descstring,_=Decode_Gen(d['op'],False)
            descstring="("+descstringl+descstring+descstringr+")"
            opstring="("+opstringl+opstring+opstringr+")"
        print("\t*** in binop",opstring,descstring)
        return opstring,descstring,descstring
    elif isinstance(node,ast.Load):
        print("Load Found")
        return 'Load','Load','Load'
    elif isinstance(node,ast.arg):
        print("Argument Found")
        return 'Arg','Arg','Arg'
    else:
        print("Got to Statement not found:",node,str_node(node))
        return '?'
    return

# Decode nodes of type Name and returns location in sheet
def Decode_Name(node,argnumflg,pytflg):
    d=dict(ast.iter_fields(node))
    if pytflg or d["id"] in ('list','range','filter','map'):
        print("Name Decoded:",d["id"])
        return d["id"]
    else:
        if argnumflg:
            print("Name Decoded, with numflag, value:",Argdict[d["id"]][1])
            return Argdict[d["id"]][1]
        else:
            print("Name Decoded, with no numflag, value:",Argdict[d["id"]][0])
            return Argdict[d["id"]][0]
# Decode nodes of type Num
def Decode_Num(node,argnumflg):
    d=dict(ast.iter_fields(node))
    if argnumflg:
        return(d['value'])
    else:
        return str(d['value'])
class RangeClass:
    def __init__(self,ref,start=0,stop=1,step=1):
        self.start=start
        self.stop=stop
        self.step=step
        self.ref=ref
    def makelist(self):
        opstring="(ROW()-ROW("+self.ref+"))*"+self.step+"+"+self.start
        opstring="IF("+opstring+"<"+self.stop+","+opstring+","+'"""")'
        print("Range Class formula =",opstring)
        return opstring

class FilterClass:
    def __init__(self,ref):
        self.ref=ref
    def makelist(self):
        refcol=chr(ord(self.ref[0])+1)
        refrow=int(self.ref[1])-1
        endrow=str(NUMFOLDS+1)
        opstring="IF("+self.ref+'="""","""",MAX('+refcol+"$"+str(refrow)+":"+refcol+str(refrow)+")+1)"
        print("Filter Class formula =",opstring)
        opstring1="IFERROR(INDEX("+self.ref[0]+"$"+self.ref[1]+":"+self.ref[0]+"$"+endrow+\
                  ",MATCH(ROW()-ROW("+chr(ord(self.ref[0])-2)+"$"+str(refrow)+"),"+refcol+"$"+\
                  self.ref[1]+":"+refcol+"$"+endrow+',0)),"""")'
        return opstring,opstring1

class MapClass:
    def __init__(self,ref):
        self.ref=ref
    def makelist(self):
        opstring=self.ref
        print("Map Class formula =",opstring)
        return opstring


    
# Visit each node of the tree in recursive fashion
def ast_visit(node, par,level=0):
    # Allow function to access global variable
    global lines,opstring
    # print out node at current level
    statement=str_node(node)
    print(lines,'|',level,'  ' * level + statement)
    lines+=1
    opstring=""
    opstring=Decode_Gen(node,False)[0]
    print("OP Statement is: ",opstring)
    # if its the root node put this in display tree as root with no parent
    if (level==0):
        disptree.create_node(node.__class__.__name__,par)
    # Traverse the Abstract Syntax Tree 
    for field, value in ast.iter_fields(node):
        if isinstance(value, list):
            for item in value:
                if isinstance(item, ast.AST):
                    prnt=item.__class__.__name__+str(lines)
                    disptree.create_node(item.__class__.__name__,prnt,parent=par)
                    ast_visit(item, prnt,level=level+1)
        elif isinstance(value, ast.AST):
            prnt=value.__class__.__name__+str(lines)
            disptree.create_node(value.__class__.__name__,prnt,parent=par)
            ast_visit(value, prnt,level=level+1)
        elif (value is not None):
            disptree.create_node(str(value),str(value)+str(lines),parent=par)
# Main Code
# Reset global line counter, each line of output will have a unique "lines" identifier
lines=0
Argdict={}      # Dictionary to hold variables, their equivalent cell references and values
Absflg=True     # Flag to make cell references absolute rather than relative
Formula=[]      # List containing formulas to be output to CSV file
Desccol=[]      # List containing column descriptors
Copydown=[]     # List containing flag for each column defining on whether it is copied down
opstring=""
testcode=readfunctions('input_file.txt')
returnexp=testcode[0][testcode[0].lower().find("return")+6:]
tree=ast.parse(testcode[0])
disptree=Tree()
ast_visit(tree,"root",0)
print()
disptree.show()
print("\nTest Code:",testcode,"\nFunction Name: ",Funcname,"\nVariables: ",Argdict,"\nFormula: ",Formula)
opsheetCSV(Funcname,Argdict,Formula,returnexp)