"""
expression_parser.py is a TatSu
provides a simplified EBNF grammer of

TODO: IF needs to be implemented as an if so unused branches don't execute.
Example: avoid the 1+#VALUE! in this:
IF(ISERROR(tk_DataValue("UK.GMPIncOrd",self.StatutoryFactorsD37)),1,1+tk_DataValue("UK.GMPIncOrd",self.StatutoryFactorsD37))
BUT: This doesn't really solve my problem, just avoids it.

By Michael Grazebrook of Joined Up Finance Ltd
"""
import tatsu
import json


__GRAMMAR = r'''
@@grammar :: ExcelExpression

# start = {term} $ ;
start = {expression}+ $
   ;

expression = term operator expression
    | term
    ;

term = const
    | function
    | group # a grouping expression
    | range
    | name
    | "None"
    ;

function = name group
    ;

# TODO: Excel, like Python, permits a trailing comma
group = '(' ','%{ expression } ')'
    ;

name = /[a-zA-Z]\w*/
    ;

range = [sheet] cell [reference_operator cell ]
    ;


reference_operator = ':' 
#    | ',' # Range union operator  e.g. "C19:C20,D19:D20"
#    | ' ' # Range intersection operator e.g. "C19:D21 D20:E22"
    ;

sheet = "'" /\w(\w|\s)*/ "'!"
    | /[a-zA-Z]\w*/ '!'
    ;

cell = /\$?[A-Z]+\$?[0-9]+/
    ;

const = number
    | text
    ;

number = /[+-]?\d+\.?\d*%?/
    ;

text = '\"' /[^"]+/ '\"' 
    ;

operator = '>=' | '<=' | '<>' | '+' | '-' | '*' | '/' | '^' | '=' | '>' | '<' | '&' 
    ;
'''

# Test terminals
text = "None"
text = "2.3 1 -3 -0.4"  # {const}+
text = '"This is some text"'  # text
text = '+-*/%^=>< >= <= <> &'  # {operator}+
text = "A1 AX23 $BB$29 $P3 P$3"  # {cell}+
text = "A4:C7 C7:f9"  # {range} - space operator in range TODO: Not working as I'm not treating spaces as significant.
text = "(E123-AVC_Fund)"

text = "MAX(ROUND((E123-AVC_Fund)/Comm_fac_at_DOC,Money_Round),0)"
text = 'YEAR(PPD_Pre06)-Sheet1!A1\n'


def expression_parser():
    return tatsu.compile(__GRAMMAR)


if __name__ == "__main__":
    class Pythonify:
        pass

    with open('../expressions.txt') as f:
        text = f.read()

    text = "tk_AnnivAfter(Date_2,DOC,)"  # Trailing commas

    parser = expression_parser()
    ast = parser.parse(text, semantics=Pythonify())
    print(json.dumps(tatsu.util.asjson(ast), indent=2))
