-- Traemos los Datos de Sega.

SELECT *
FROM OPENROWSET('SQLNCLI','Server=CLMOSVRSEGA01P;UID=usuriesgo;PWD=rsgo.seg14;',
'select *
from [BD_Riesgo].dbo.Piloto_Aprobado_CF')

SELECT *
FROM OPENROWSET('SQLNCLI','Server=CLLOGDR02P;Trusted_Connection=Yes;Initial Catalog=pubs;Integrated Security=SSPI;',
'select distinct *, cast(substring(g_cliente,1,len(g_cliente)-1) as int) as rut
from [CLLOGDR02P].tempdb..##mc_tabla_para_arayi')

