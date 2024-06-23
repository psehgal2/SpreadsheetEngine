import context
import sheets.Parser as Parser
import pytest
import lark
import random
from sheets.CellError import CellError
from sheets.CellErrorType import *
import decimal

def test_simple_addition():
   parser = lark.Lark.open('sheets/formulas.lark', start='formula')
   ITERS = 100
   evaluator = Parser.FormulaEvaluator(None,"",{})
   for i in range(ITERS):
        l = decimal.Decimal(random.random())
        op = random.choice(['+','-'])
        r = decimal.Decimal(random.random())
        test_str = "=" + str(l) + op + str(r)  
        parsed = parser.parse(test_str)
        result = evaluator.visit(parsed)
        if op == '-':
            assert (result == (l - r))
        if op == '+':
            assert (result == (l + r))

def test_simple_multiplication():
   parser = lark.Lark.open('sheets/formulas.lark', start='formula')
   ITERS = 100
   evaluator = Parser.FormulaEvaluator(None,"",{})
   for i in range(ITERS):
        l = decimal.Decimal(random.random())
        op = random.choice(['*','/'])
        r = decimal.Decimal(random.random())
        test_str = "=" + str(l) + op + str(r)
        parsed = parser.parse(test_str)
        result = evaluator.visit(parsed)
        if op == '/':
            assert (result == (l / r))
        if op == '*':
            assert (result == (l * r))

def test_simple_unary():
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    ITERS = 100
    evaluator = Parser.FormulaEvaluator(None, "",{})
    for i in range(ITERS):
        op = random.choice(['+', '-'])
        r = decimal.Decimal(random.random())
        test_str = "=" + op +  str(r)
        parsed = parser.parse(test_str)
        result = evaluator.visit(parsed)
        if op == '+':
            assert(result == r)
        elif op == '-':
            assert(result == -r)

class ParserTestHelper():
    def __init__(self,parser, evaluator):
        self.parser = parser
        self.evaluator = evaluator
    def check(self,formula,value):
        formula = "=" + formula
        parsed = self.parser.parse(formula)
        result = self.evaluator.visit(parsed)
        assert (value == result)



# def test_simple_expression():
#     parser = lark.Lark.open('sheets/formulas.lark', start='formula')
#     ITERS = 100
#     evaluator = Parser.FormulaEvaluator(None,"")
#     helper = ParserTestHelper(parser,evaluator)
#     helper.check("6*9 - (4/2) + 5.6",decimal.Decimal(6*9 - (4/2) + 5.6))
#     helper.check("(100 * (-1) + 5.6 * 32 - 5 * 6)/32 + 1)",decimal.Decimal((100 * (-1) + 5.6 * 32 - 5 * 6)/32 + 1))
#     helper.check("36/22 - 1 + 5/(2+1) * 5 - (4) + 42)", decimal.Decimal(36/22 - 1 + 5/(2+1) * 5 - (4) + 42))
#     helper.check("5/5/5/5/5*(5+1)/3 + 5 +1 *(-1)", decimal.Decimal(5/5/5/5/5*(5+1)/3 + 5 +1 *(-1)))


def test_concatentation():
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    ITERS = 100
    evaluator = Parser.FormulaEvaluator(None,"", {})

    result = evaluator.visit(parser.parse("=(\"Hello\") & (\"jim\")"))
    assert(result== "Hellojim")
    
def test_error():
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    ITERS = 100
    evaluator = Parser.FormulaEvaluator(None,"",{})

    result = evaluator.visit(parser.parse("=(5*2 + 9/0)"))
    assert isinstance(result,CellError)
    result = evaluator.visit(parser.parse("=(5*2 + 9) * #REF!"))
    assert isinstance(result,CellError)
    result = evaluator.visit(parser.parse("=-#REF!"))
    assert isinstance(result,CellError)



def test_error_presedence():
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    ITERS = 100
    evaluator = Parser.FormulaEvaluator(None,"",{})

    result = evaluator.visit(parser.parse("=(#REF!) * (#CIRCREF!) + 5 / 2 - 6"))
    assert isinstance(result,CellError)
    assert (result.get_type() == CellErrorType.CIRCULAR_REFERENCE)
    result = evaluator.visit(parser.parse("=(5*2 + 9) * #REF! +  (#CIRCREF! - 1)  / (#NAME?)"))
    assert isinstance(result,CellError)
    assert (result.get_type() == CellErrorType.CIRCULAR_REFERENCE)
    result = evaluator.visit(parser.parse("=(#ERROR!/2 + 1) - #REF! * 2 / #DIV/0! "))
    assert isinstance(result,CellError)
    assert (result.get_type() == CellErrorType.PARSE_ERROR)
    result = evaluator.visit(parser.parse("=(#ERROR! * #CIRCREF!)/2 - 5/ (#DIV/0!) + 1000 "))
    assert isinstance(result,CellError)
    assert (result.get_type() == CellErrorType.PARSE_ERROR)
    result = evaluator.visit(parser.parse("=(#DIV/0! * 100) "))
    assert isinstance(result,CellError)
    assert (result.get_type() == CellErrorType.DIVIDE_BY_ZERO)

# # def test_cell_expression():
#     parser = lark.Lark.open('sheets/formulas.lark', start='formula')
#     ITERS = 100
#     evaluator = Parser.FormulaEvaluator(None,"")
#     parsed = parser.parse('=3*A2')
#     evaluated = evaluator.evaluate(parsed)



