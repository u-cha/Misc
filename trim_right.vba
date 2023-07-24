
Option Explicit


        'this function is designed to trim unwanted spaces(or other redundant symbols)
        'to the right from your string and it is not meant to do anything else or to
        'be bulletproof from what you actually input.
        'Sample syntax is TRIM_RIGHT(A2; " ")


        Function TRIM_RIGHT(cell_with_text, what_we_trim):
        Dim text_length As Integer
        Dim p As Integer


        text_length = Len(cell_with_text)
        p = text_length
        While Mid(cell_with_text, p, 1) = what_we_trim
        p = p - 1
        Wend


        TRIM_RIGHT = Mid(cell_with_text, 1, p)




        End Function