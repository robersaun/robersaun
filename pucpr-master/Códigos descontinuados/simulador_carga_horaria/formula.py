letiva = 20
nao_letiva = 20
campus = "Curitiba"
total = letiva + nao_letiva
porcentagem_nao_letiva = nao_letiva / total

print(letiva, nao_letiva, campus, total, porcentagem_nao_letiva)

if total > 40:
    # G2>40
    print("Excedeu CH")
elif total == 40 and porcentagem_nao_letiva >= 0.5:
    # =E([@Total]=40;[@[Horas Não Letivas]]/[@Total]>=0,5)
    print("TI")
elif campus == "Curitiba" and total >= 36 and porcentagem_nao_letiva >= 0.5:
    # E([@Total]>=36;[@[Horas Não Letivas]]/[@Total]>=0,5;[@Campus]="Curitiba");"TI"
    print("TI")
elif campus == "Toledo" and total >= 36 and porcentagem_nao_letiva >= 0.5:
    # E([@Total]>=36;[@[Horas Não Letivas]]/[@Total]>=0,5;[@Campus]="Toledo")
    print("TI")
elif total >= 12 and porcentagem_nao_letiva >= 0.25:
    print("TP")
else:
    print("Horista")
