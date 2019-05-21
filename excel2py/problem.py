import tatsu

grammar = r'''
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

group = '(' ','%{ expression } ')'
    ;

name = /[a-zA-Z]\w*/
    ;

range = [sheet] cell [reference_operator cell ]
    ;

reference_operator = ':' | ',' | ' '
    ;

sheet = "'" /\w(\w|\s)*/ "'!"
    | name '!'
    ;

cell = /\$?[A-Z]+\$?[0-9]+/
    ;

const = number
    | text
    ;

# negation = '-' ;
# 
# number = positive
#     | negation positive ;
number = /[+-]?\d+\.?\d*%?/
    ;

positive = /\d+\.?\d*%?/ ;

# TOOD: Should negation be separate?
#_number = /[+-]?\d+\.?\d*%?/
#    ;

text = '\"' /[^"]+/ '\"' 
    ;

operator = '>=' | '<=' | '<>' | '+' | '-' | '*' | '/' | '^' | '=' | '>' | '<' | '&' 
    ;
'''

text = "None"
text = "2.3 1 -3 -0.4"  # {number}+

text = "MAX(ROUND((E123-var1)/var2,money_round),0)"


class Pythonify:
    # def start(ast):
    #     print(f"start: {ast}")
    #     return flatten(ast)
    #
    # def function(ast):
    #     print(f"function: {ast}")
    #     return Pythonify._default(ast)
    #
    # def term(ast):
    #     print(f"term: {ast}")
    #     return ast
    #
    # def group(ast):
    #     print(f"group: {ast}")
    #     return ast
    #
    # def operator(ast):
    #     print(f"operator: {ast}")
    #     return Pythonify._default(ast)

    # def expression(ast):
    #     print(f"expression: {ast}")
    #     return ast
    #
    # def number(ast):
    #     print(f"number: {ast}")
    #     return Pythonify._default(ast)

    def _default(ast):
        return flatten(ast)


def flatten(ast):
    if isinstance(ast, str):
        return ast
    if isinstance(ast, list):
        return ''.join([flatten(bit) for bit in ast])
    assert False, str(ast)


def run():
    with open('../tests/expressions.txt') as f:
        lines = f.readlines()

    # lines = [ 'YEAR(Some_year)-1\n']

    pythonifier = tatsu.compile(grammar)

    for line_number, line in enumerate(lines):
        line = line.strip()
        ast = pythonifier.parse(line, semantics=Pythonify)
        if line != ast:
            print(f"\n{line_number:-3}: {line}")
            print("    ", ast)


run()
