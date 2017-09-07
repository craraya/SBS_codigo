

SELECT *
into [RT_ANALYTICS_ORI].dbo.CA_20161130_Piloto_Aprobado_CF
FROM OPENROWSET('SQLNCLI','Server=CLMOSVRSEGA01P;UID=usuriesgo;PWD=rsgo.seg14;',
'select *
from [BD_Riesgo].dbo.Piloto_Aprobado_CF')

