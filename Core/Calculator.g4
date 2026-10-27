grammar Calculator;

compileUnit
    : expression EOF
    ;

expression
    : NUMBER                                                                # NumberExpr
    | IDENTIFIER                                                            # IdentifierExpr
    | '(' expression ')'                                                    # ParenthesesExpr
    | operator=('+' | '-') expression                                       # UnaryExpr
    | <assoc=right> left=expression '^' right=expression                    # PowerExpr
    | left=expression operator=('*' | '/' | 'mod' | 'div') right=expression # MultiplicativeExpr
    | left=expression operator=('+' | '-') right=expression                 # AdditiveExpr
    ;

NUMBER     : [0-9]+ ('.' [0-9]+)? ;
IDENTIFIER : [A-Za-z]+ [0-9]+ ;
WS         : [ \t\r\n]+ -> skip ;