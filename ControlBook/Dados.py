class Dados:

    def __init__(self):
        self.colunas = ['Número', 'Categoria', 'Origem de Abertura', 'Serviço (Completo)', 'Aberto em', 'Horário do Incidente',
                        'Data de Normalização', 'Resolvido em', 'Fechado em', 'Criado por', 'Serviço (1º Nível)', 'Assunto',
                        'Descrição', 'Cliente (Organização)', 'Unidade', 'Responsável: Equipe', 'Causa Raiz', 'Resolução',
                        'Indicador do SLA de Solução',
                        'Tempo parado (Horas corridas)', 'Tempo de vida (Descontando parado - Horas corridas)',
                        'Tempo de vida (Horas corridas)',
                        'Ticket Externo']

        self.dados = ['ToIP', 'Solicitações', 'Vídeo', 'Dados', '»', 'Incidentes', 'Suporte Interno','Administração dos Ambientes']

        self.gtw_sebrae = {'JUAZEIRO':'GW | EP_GWFXO_JUAZEI',
                           'ITAPIPOCA':'GW | EP_GWFXO_ITAPIP',
                           'IGUATU':'GW | EP_GWFXO_IGUATU',
                           'TAUA':'GW | EP_GWFXO_TAUA',
                           'CRATEUS':'GW | EP_GWFXO_CRATEUS',
                           'SOBRAL':'GW | EP_GWFXO_SOBRAL',
                           'QUIXERAMOBIM':'GW | EP_GWFXO_QUIXERAMOBIM',
                           'LIMOEIRO':'GW | EP_GWFXO_LIMOEIRO',
                           'BATURI':'GW | EP_GWFXO_BATURI',
                           'ARACATI':'GW | EP_GWFXO_ARACAT'}

        self.dados_gtw_sebrae = ['ITAPIPOCA','IGUATU','TAUA','CRATEUS','SOBRAL','QUIXERAMOBIM','LIMOEIRO','BATURI','ARACATI','JUAZEIRO']


        self.clientes_unidades = ['UNI_ACUJT','UNI_ACUMT','UNI_ACUNO','UNI_ACUPQ','UNI_ACURE','UNI_ACUTO','UNI_AREC','UNI_UCUEB',
                    'LJSAC','LJRES','LJJFO','LJCGC','LJCAS','LJANT','CMCPN','CMANA-CMDAN','CDSSA','AEXUDI','AEXMON',
                    'ACZIM','ACZGP','ACZGG','ACZFT','ACZBT','ACVVA','ACVRE','ACVOP','ACVNA','ACVHD','ACVCA','ACURN',
                     'ACUBT','ACTUR','ACTSR','ACTIJ','ACTHE','ACTBO','ACTAN','ACSST','ACSPM','ACSOB','ACSMA','ACSLZ',
                     'ACSLP','ACSJR','ACSJQ','ACSHN','ACSER','ACSCZ','ACSCS','ACSAD','ACSAC','ACSAB','ACRJF','ACRIA',
                     'ACREZ','ACREC','ACRCR','ACRBP','ACRBC','ACPRM','ACPPD','ACPOV','ACPOA','ACPNI','ACPIT','ACPIN',
                     'ACPGO','ACPGA','ACPEV','ACPEH','ACPDP','ACPCS','ACPBA','ACOUB','ACOSC','ACORI','ACNOH','ACNIT',
                     'ACNIG','ACNAT','ACMSO','ACMON','ACMNA','ACMIL','ACMII','ACMGF','ACMCZ','ACMAU','ACMAE','ACMAC',
                     'ACMAB','ACLPT','ACLPA','ACLME','ACLLI','ACLJE','ACLIN','ACLDB','ACLAU','ACJZN','ACJQE','ACJPA',
                     'ACJQE','ACJPA','ACJOI','ACJMO','ACJLE','ACJKK','ACJIP','ACJER','ACJBT','ACJAU','ACJAR','ACJAN',
                     'ACIUN','ACITP','ACIPR','ACIPG','ACIMP','ACIBR','ACIBP','ACIBO','ACHSA','ACHBU','ACGVA','ACGUA',
                     'ACGRU','ACGRA','ACGOI','ACGIG','ACFSA','ACFRA','ACFOP','ACFER','ACFAL','ACDVN','ACDEL','ACDDA',
                     'ACDCX','ACCVI','ACCVG','ACCUI','ACCTP','ACCTI','ACCSU','ACCST','ACCSP','ACCSM','ACCSD','ACCSM',
                     'ACCSD','ACCSA','ACCOL','ACCNV','ACCNS','ACCNO','ACCNN','ACCJZ','ACCIN','ACCGO','ACCDX','ACCCV',
                     'ACCCV','ACCCO','ACCCN','ACCBT','ACCBL','ACCAV','ACCAU','ACCAP','ACCAM','ACCAH','ACCAE','ACCAB',
                     'ACBVT','ACBUT','ACBTF','ACBTA','ACBRJ','ACBRG','ACBRF','ACBON','ACBEF','ACBDP','ACBDO','ACBCA',
                     'ACBAU','ACBAR','ACAVA','ACATU','ACATM','ACARA','ACAQM','ACAQA','ACAPI','ACANG','ACANA','ACALP',
                     'AATHE','AASLZ','AASDU','AARNV','AARIA','AAREC','AAPPD','AAPOV','AAPOA','AAPET','AANAT-BOX','AAMON',
                     'AAMII','AAMGF','AAMCZ-BOX','AAMCZ-AEX','AAMAU','AAMAC','AAMAB','AALDB','AAJZN','AAJAP','AAJER',
                     'AAIMP','AAGRU','AAGRU_BOX','AAGIB','AAFOZ','AAFOR','AACWB-AEX','AACPQ','AACKS','AABVT','AABEL_CONTAINER',
                     'AABEL','AABAU_BOX','AAATM','AAAGRU_BOX','PAÇO','ZOONOSES','CD CIAL','ARGOS','ALTO CUSTO','VIGILANCIA SANITARIA',
                     'TRANSITO','FEPASA','HOLANDA','JOINVILLE','SARZEDO','FAB4','CDC','HACASA','CANADA','UNI_ACUJG',
                     'UNI_ACUJT','UNI_ACUIA','UNI_ACUFU','UNI_ACUBP','UNI_ACUBO','ACUBH']

