def mostrar_empresa ():
        
        print("ASR ELEVADORES")
        print("Sistema de orçamentos")
        print(1.0)
mostrar_empresa()

nome = input("nome do cliente:")

produto = input("Produto:")

if produto == "andaime_suspenso":
    preco_unitario = 180000

elif produto == "grua_torre":
    preco_unitario = 250000

else:
    print("Produto não encontrado")
    exit()

quantidade = int(input("Quantidade: "))


  #parcelando o pagamento
valor_total = preco_unitario * quantidade

parcela1 = valor_total * 0.50
parcela2 = valor_total * 0.15
parcela3 = valor_total * 0.15
parcela4 = valor_total * 0.10
parcela5 = valor_total * 0.05
parcela6 = valor_total * 0.05

print("Cliente:", nome)
print("Produto:", produto)
print("Valor total:", valor_total)
print("Valor por parcela:", parcela1)
def calcular_parcelas(valor_total):
    parcela1 = valor_total * 0.50
    parcela2 = valor_total * 0.15
    parcela3 = valor_total * 0.15
    parcela4 = valor_total * 0.10
    parcela5 = valor_total * 0.05
    parcela6 = valor_total * 0.05

    print("Parcela 1:", parcela1)
    print("Parcela 2:", parcela2)
    print("Parcela 3:", parcela3)
    print("Parcela 4:", parcela4)
    print("Parcela 5:", parcela5)
    print("Parcela 6:", parcela6)

calcular_parcelas(valor_total)
def calcular_total(preco_unitario, quantidade):
     return preco_unitario * quantidade
valor_total = calcular_total(preco_unitario, quantidade)
calcular_total(preco_unitario, quantidade)