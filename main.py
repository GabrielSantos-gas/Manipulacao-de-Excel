from datetime import date
from openpyxl.chart import  Reference
from classes import LeitorAcoes, GerenciadorPlanilha, PropriedadeSeriesGrafico
from openpyxl.styles import Font, PatternFill, Alignment


try:

    acao = "BIDI4"

    leitor_acoes = LeitorAcoes(caminho_arquivo='./dados/')
    leitor_acoes.processa_arquivo(acao)

    gerenciador = GerenciadorPlanilha()
    planilha_dados = gerenciador.adiciona_planilha("Dados")

    gerenciador.adiciona_linha(["DATA", "COTAÇÃO", " BANDA INFERIOR", "BANDA SUPERIOR"])

    indice = 2

    for linha in leitor_acoes.dados:
        # data
        ano_mes_dia = linha[0].split(" ")[0]
        data = date(
            year=int(ano_mes_dia.split("-")[0]),
            month=int(ano_mes_dia.split("-")[1]),
            day=int(ano_mes_dia.split("-")[2]))

        # cotacao
        cotacao = float(linha[1])

        formula_bb_inferior = f'=AVERAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})'
        formula_bb_superior = f'=AVERAGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})'
        # atualiza as celular da planilha do excel
        gerenciador.atualiza_celula(celula=f'A{indice}', dado=data)
        gerenciador.atualiza_celula(celula=f'B{indice}', dado=cotacao)
        gerenciador.atualiza_celula(celula=f'C{indice}', dado=formula_bb_inferior)
        gerenciador.atualiza_celula(celula=f'D{indice}', dado=formula_bb_superior)

        indice += 1

    gerenciador.adiciona_planilha(titulo_planilha='Gráfico')

    # mesclagem de celula para criação do cabeçalho do Gráfico
    gerenciador.mescla_celular(celula_inicio='A1', celula_fim='T2')

    gerenciador.aplica_estilos(
        celula='A1',
        estilos=[
            ('font', Font(b=True, sz=18, color='FFFFFF')),
            ('alignment', Alignment(vertical="center", horizontal="center")),
            ('fills', PatternFill("solid", fgColor='07838f')),
        ]
    )

    gerenciador.atualiza_celula('A1', "Histórico de Cotações")

    referencia_cotacoes = Reference(planilha_dados, min_col=2, min_row=2, max_col=4, max_row=indice)
    referencia_datas = Reference(planilha_dados, min_col=1, min_row=2, max_col=1, max_row=indice)

    gerenciador.adicona_grafico_linha(
        celula='A3',
        comprimento=30,
        altura=10,
        titulo=f'Cotações - {acao}',
        titulo_eixo_x="data da Cotação",
        titulo_eixo_y="valor da Cotação",
        referencia_eixo_y=referencia_datas,
        referencia_eixo_x=referencia_cotacoes,
        propriedades_grafico=[
            PropriedadeSeriesGrafico(grossura=0, cor_preenchimento='0a55ab'),
            PropriedadeSeriesGrafico(grossura=0, cor_preenchimento='0a5500'),
            PropriedadeSeriesGrafico(grossura=0, cor_preenchimento='12a154'),
        ]
    )

    gerenciador.mescla_celular(celula_inicio='G23', celula_fim='J27')
    gerenciador.adiciona_imagem(celula='G23', caminho_imagem="./recursos/logo.png")

    gerenciador.salva_arquivo("./saida/PlanilhaRefaturada.xlsx")

except AttributeError:
    print('Atributo inexistente.')

except ValueError:
    print('Formato de dados incorretos!')

except FileNotFoundError:
    print('Arquivo nao encontrado')


except Exception as excecao :
    print(f'Ocorreu um erro na execução do programa. Erro: {str(excecao)}')