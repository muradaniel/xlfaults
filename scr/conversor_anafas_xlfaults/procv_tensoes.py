def ProcV_tensoes(Barra, Linha, Maquina, Transformador, Carga):
    
    Linha['Tensão (kV)'] = Linha['Barra de'].map(Barra.set_index('Número')['Tensão (kV)'])

    Maquina['Tensão (kV)'] = Maquina['Barra Conectada'].map(Barra.set_index('Número')['Tensão (kV)'])

    Transformador['Tensão Primário (kV)'] = Transformador['Barra de'].map(Barra.set_index('Número')['Tensão (kV)'])
    Transformador['Tensão Secundário (kV)'] = Transformador['Barra para'].map(Barra.set_index('Número')['Tensão (kV)'])
    
    return Linha, Maquina, Transformador, Carga 