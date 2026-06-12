# ruff: noqa
def gerar_assunto_email() -> str:
    return 'Monitoramento de remessas Pendentes da Rede Credenciada - Competências 2022 a 2025'


def gerar_corpo_email(regional: str, nome_anexo_pdf: str) -> str:
    return f"""
		À Unidade Regional de {regional},
		
		
		A Gerência de Auditoria e Contas da Saúde (GEACS) está realizando o monitoramento das remessas de faturamento da Rede Credenciada que permanecem pendentes de conclusão, com o objetivo de promover, em conjunto com as Unidades Regionais, a regularização dos processos e a adequada efetivação das etapas de auditoria, processamento e pagamento.
		
		Nesse sentido, encaminhamos em anexo a relação consolidada das remessas das competências de 2022 a 2025 que permanecem com status "Em Revisão", "Liberado para Pagamento" e "Pagamento Processado" com valor a pagar igual a zero, para conferência, adoção das providências cabíveis e retorno quanto ao andamento das tratativas.
		
		Para auxiliar a análise das remessas relacionadas, encaminhamos também o documento "{nome_anexo_pdf}", contendo orientações específicas para tratamento de cada situação identificada na planilha.
		
		Ressaltamos que, conforme disposto nos normativos e fluxos institucionais vigentes, compete às Unidades Regionais o acompanhamento do faturamento dos prestadores credenciados sob seu polo de atuação, incluindo o monitoramento das etapas do processo e a adoção das medidas necessárias para sua conclusão tempestiva.
		
		Considerando o período das competências em análise e a necessidade de mitigação de riscos relacionados à prescrição de despesas, solicitamos especial atenção à conferência e regularização das remessas encaminhadas, priorizando-se, sempre que possível, a análise em ordem cronológica das competências mais antigas.
		
		A GEACS acompanhará periodicamente a evolução das remessas relacionadas e poderá solicitar informações complementares sobre as providências adotadas e eventuais dificuldades identificadas durante a regularização dos processos.
		
		As dúvidas técnicas relacionadas às remessas em revisão poderão ser direcionadas aos respectivos departamentos de auditoria responsáveis pelo atendimento da regional.
		
		Contamos mais uma vez com o apoio e colaboração dessa Unidade Regional para a conclusão das pendências identificadas.
		
		Permanecemos à disposição para esclarecimentos, que poderão ser encaminhados pelo Teams à servidora Rayane Paula Arcendino de Sousa, pelo telefone (31) 3915-2149 ou por meio deste e-mail.
		
		
		Atenciosamente,
		"""
