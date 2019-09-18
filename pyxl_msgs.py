class msgs:
    def __init__(self):
        self.menu = """
////////////////////////////////////
//         MENU DE OPÇÕES         //
//      *** NÃO HABILITADOS       //        
////////////////////////////////////
// A - Alimentar Controle         //
// B - Procurar UL-CCT-IP         //
// C -                        *** //
// D -                        *** //
// E - Planilhas utilizadas       //
// Q - Finalizar Programa         //
////////////////////////////////////
"""
        self.msg1 = '''
                OPÇÃO INCORRETA!'''

    def getMenu(self):
        return print(self.menu)

    def getMsg1(self):
        return print(self.msg1)