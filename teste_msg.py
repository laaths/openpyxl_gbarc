class msgs:
    def __init__(self):
        self.menu = """
////////////////////////////////////
//     MENU DE OPÇÕES             //
////////////////////////////////////
// A - Copy/Paste 2 planilhas     //
// B - Procurar UL-CCT-IP-CNPJ    //
// C - Inserir Loopback           //
// D - Atividade lotericos        //
// E - //
// Q - Finalizar Programa         //
////////////////////////////////////
"""
        self.msg1 = '''
                OPÇÃO INCORRETA!'''

    def getMenu(self):
        return print(self.menu)

    def getMsg1(self):
        return print(self.msg1)