from classes import cadastrar, arquivos

planilha = 'nova.xlsx' # Digita o nome da planilha sendo utilizada
aba_plan = 'retiradas' # Nome da aba dentro do arquivo que vai utilizar

cad = cadastrar(input("Nome: ").capitalize(), input("Nº Serie: "))
read = arquivos()
wb = read.rArq(planilha, aba_plan)
aba = wb[aba_plan]
ws = wb.active

cad.cadastramento(aba, ws)

wb.save(planilha)
