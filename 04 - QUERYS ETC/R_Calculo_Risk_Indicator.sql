

UPDATE #CALCULO_RI 
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente <516                                            THEN '000-515'
                                     WHEN ScoreSinacofiCliente >=516 AND ScoreSinacofiCliente <685 THEN '516-684'
                                     WHEN ScoreSinacofiCliente >=685 AND ScoreSinacofiCliente <771 THEN '685-770'
                                     WHEN ScoreSinacofiCliente >=771 AND ScoreSinacofiCliente <804 THEN '771-803'
                                     WHEN ScoreSinacofiCliente >=804                                            THEN '804-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE       C_SIS_LABEL='PB' AND    N_LOGICPATH_LABEL='Nuevo Campana' -- AND CATEG_SINACOFI IS NULL ORDER BY           ScoreSinacofiCliente      

            -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI



 UPDATE #CALCULO_RI 
 SET CATEG_SCORE_INT= CASE WHEN SCORE_PB_ORI_CALCULADO <598                                                     THEN '000-597'
                                     WHEN SCORE_PB_ORI_CALCULADO >=598 AND SCORE_PB_ORI_CALCULADO <609 THEN '598-608'
                                     WHEN SCORE_PB_ORI_CALCULADO >=609 AND SCORE_PB_ORI_CALCULADO <621 THEN '609-620'
                                     WHEN SCORE_PB_ORI_CALCULADO >=621 AND SCORE_PB_ORI_CALCULADO <655 THEN '621-654'
                                     WHEN SCORE_PB_ORI_CALCULADO >=655                                               THEN '655-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE       C_SIS_LABEL='PB' AND    N_LOGICPATH_LABEL='Nuevo Campana'
--AND CATEG_SINACOFI IS NULL 
 --ORDER BY SCORE_PB_ORI_CALCULADO   
 

  
   -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI

UPDATE #CALCULO_RI 
 SET RI= CASE                WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
                                                                                                          
                                    WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
                                    WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
                                                                                                          
                                    WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='621-654'  THEN 'B'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
                                                                                                          
                                    WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='000-597'  THEN 'B'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='598-608'  THEN 'B'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='609-620'  THEN 'B'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
                                                                                                          
                                    WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='000-597'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='598-608'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='609-620'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
                                                                      END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
                                    WHERE      C_SIS_LABEL='PB' AND       N_LOGICPATH_LABEL='Nuevo Campana'
--AND RI IS NULL 
 --ORDER BY RI      
 
  -- SELECT distinct C_SIS_LABEL,N_LOGICPATH_LABEL FROM RT_SCORING.DBO.CALCULO_RI WHERE       C_SIS_LABEL='PB' AND    N_LOGICPATH_LABEL='Nuevo Campana'
                                    
---------------------------                      
 /* 2  PB NUEVO NO CAMPAÑA */      --    WHERE       C_SIS_LABEL='PB' AND       N_LOGICPATH_LABEL='Nuevo'            
--------------------------- 
            
            
 UPDATE #CALCULO_RI 
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente <516                                          THEN '000-515'
                                     WHEN ScoreSinacofiCliente >=516 AND ScoreSinacofiCliente <685 THEN '516-684'
                                     WHEN ScoreSinacofiCliente >=685 AND ScoreSinacofiCliente <771 THEN '685-770'
                                     WHEN ScoreSinacofiCliente >=771 AND ScoreSinacofiCliente <804 THEN '771-803'
                                     WHEN ScoreSinacofiCliente >=804                                          THEN '804-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE       C_SIS_LABEL='PB' AND    N_LOGICPATH_LABEL='Nuevo'    -- AND CATEG_SINACOFI IS NULL ORDER BY           ScoreSinacofiCliente      

            -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI


UPDATE #CALCULO_RI 
 SET CATEG_SCORE_INT= CASE WHEN SCORE_PB_ORI_CALCULADO <598                                                     THEN '000-597'
                                     WHEN SCORE_PB_ORI_CALCULADO >=598 AND SCORE_PB_ORI_CALCULADO <609 THEN '598-608'
                                     WHEN SCORE_PB_ORI_CALCULADO >=609 AND SCORE_PB_ORI_CALCULADO <621 THEN '609-620'
                                     WHEN SCORE_PB_ORI_CALCULADO >=621 AND SCORE_PB_ORI_CALCULADO <655 THEN '621-654'
                                     WHEN SCORE_PB_ORI_CALCULADO >=655                                               THEN '655-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE       C_SIS_LABEL='PB' AND    N_LOGICPATH_LABEL='Nuevo'    
--AND CATEG_SINACOFI IS NULL 
 --ORDER BY SCORE_PB_ORI_CALCULADO   
 

  
   -- SELECT DISTINCT C_SIS_LABEL,N_LOGICPATH_LABEL ,RI, CATEG_SINACOFI, CATEG_SCORE_INT FROM RT_SCORING.DBO.CALCULO_RI 

UPDATE #CALCULO_RI 
 SET RI= CASE                WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
                                   WHEN CATEG_SINACOFI='000-515'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
                                                                                                          
                                    WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='621-654'  THEN 'C'
                                   WHEN CATEG_SINACOFI='516-684'  AND CATEG_SCORE_INT='655-999'  THEN 'B'
                                                                                                           
                                    WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='000-597'  THEN 'E'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='598-608'  THEN 'D'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='609-620'  THEN 'C'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='621-654'  THEN 'B'
                                   WHEN CATEG_SINACOFI='685-770'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
                                                                                                          
                                    WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='000-597'  THEN 'B'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='598-608'  THEN 'B'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='609-620'  THEN 'B'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
                                   WHEN CATEG_SINACOFI='771-803'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
                                                                                                          
                                    WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='000-597'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='598-608'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='609-620'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='621-654'  THEN 'A'
                                   WHEN CATEG_SINACOFI='804-999'  AND CATEG_SCORE_INT='655-999'  THEN 'A'
                                                                      END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
                                   WHERE       C_SIS_LABEL='PB' AND       N_LOGICPATH_LABEL='Nuevo'    
--AND RI IS NULL 
 --ORDER BY RI      
                  
--------------------------- 
 /* 3  PB  ANTIGUO */   --    WHERE       C_SIS_LABEL='PB' AND    N_LOGICPATH_LABEL IN ('Antiguo','Antiguo Campana')
--------------------------- 

UPDATE #CALCULO_RI 
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente <381 THEN '000-380'
                                     WHEN ScoreSinacofiCliente >=381 AND ScoreSinacofiCliente <785 THEN '381-784'
                                     WHEN ScoreSinacofiCliente >=785 AND ScoreSinacofiCliente <836 THEN '785-835'
                                     WHEN ScoreSinacofiCliente >=836 AND ScoreSinacofiCliente <862 THEN '836-861'
                                     WHEN ScoreSinacofiCliente >=862 THEN '862-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE C_SIS_LABEL='PB' AND   N_LOGICPATH_LABEL IN ('Antiguo','Antiguo Campana') -- AND CATEG_SINACOFI IS NULL ORDER BY          ScoreSinacofiCliente      

            -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI
                                                                                                             
 UPDATE #CALCULO_RI                                                                            
 SET CATEG_SCORE_INT= CASE WHEN BHV_CLI_ANT <568                             THEN '000-567'                                   
                                     WHEN BHV_CLI_ANT >=568 AND BHV_CLI_ANT <623 THEN '568-622'                                                  
                                     WHEN BHV_CLI_ANT >=623 AND BHV_CLI_ANT <655 THEN '623-654'                                     
                                     WHEN BHV_CLI_ANT >=655 AND BHV_CLI_ANT <664 THEN '655-663'                                     
                                     WHEN BHV_CLI_ANT >=664                               THEN '664-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE C_SIS_LABEL='PB' AND   N_LOGICPATH_LABEL IN ('Antiguo','Antiguo Campana')                               
 --AND CATEG_SINACOFI IS NULL                                                                                
 --ORDER BY SCORE_PB_ORI_CALCULADO                                                                          
                                                                                                             
   -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI                                                                
                                                                                                             
 UPDATE #CALCULO_RI                                                                            
 SET RI= CASE            WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='000-567'  THEN 'E'                    
                                    WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='568-622'  THEN 'D'                          
                                    WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='623-654'  THEN 'C'                          
                                    WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='655-663'  THEN 'B'                          
                                    WHEN CATEG_SINACOFI='000-380'  AND CATEG_SCORE_INT='664-999'  THEN 'B'                          
                                                                                                                                    
                                    WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='000-567'  THEN 'D'                          
                                    WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='568-622'  THEN 'D'                          
                                    WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='623-654'  THEN 'C'                          
                                    WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='655-663'  THEN 'B'                          
                                    WHEN CATEG_SINACOFI='381-784'  AND CATEG_SCORE_INT='664-999'  THEN 'B'                          
                                                                                                                                    
                                    WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='000-567'  THEN 'D'                          
                                    WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='568-622'  THEN 'B'                          
                                    WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='623-654'  THEN 'A'                          
                                     WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='655-663'  THEN 'A'                          
                                    WHEN CATEG_SINACOFI='785-835'  AND CATEG_SCORE_INT='664-999'  THEN 'A'                          
                                                                                                                                    
                                    WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='000-567'  THEN 'A'                           
                                    WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='568-622'  THEN 'A'                           
                                    WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='623-654'  THEN 'A'                          
                                    WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='655-663'  THEN 'A'                          
                                    WHEN CATEG_SINACOFI='836-861'  AND CATEG_SCORE_INT='664-999'  THEN 'A'                          
                                                                                                                                    
                                    WHEN CATEG_SINACOFI='862-999' AND CATEG_SCORE_INT='000-567'  THEN 'A'                           
                                    WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='568-622'  THEN 'A'                          
                                    WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='623-654'  THEN 'A'                          
                                    WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='655-663'  THEN 'A'                           
                                    WHEN CATEG_SINACOFI='862-999'  AND CATEG_SCORE_INT='664-999'  THEN 'A'                          
                                                                       END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI                                       
                                    WHERE C_SIS_LABEL='PB' AND N_LOGICPATH_LABEL IN ('Antiguo','Antiguo Campana')                 
 --AND RI IS NULL                                                                                            
 --ORDER BY RI                                                                                                 
                                                                  
  -- SELECT distinct C_SIS_LABEL,N_LOGICPATH_LABEL FROM RT_SCORING.DBO.CALCULO_RI WHERE       C_SIS_LABEL='CF' AND    N_LOGICPATH_LABEL='Nuevo Campana'
                                    

--------------------------- 
 /* 4  CF  NUEVO CAMPAÑA  */ --    WHERE       C_SIS_LABEL='CF' AND    N_LOGICPATH_LABEL IN ('Nuevo Campana')
--------------------------- 

UPDATE #CALCULO_RI 
 SET CATEG_SINACOFI= CASE WHEN ScoreSinacofiCliente  <348                                           THEN '000-347'
                                     WHEN ScoreSinacofiCliente >=348 AND ScoreSinacofiCliente <401 THEN '348-400'
                                     WHEN ScoreSinacofiCliente >=401 AND ScoreSinacofiCliente <535 THEN '401-534'
                                     WHEN ScoreSinacofiCliente >=535 AND ScoreSinacofiCliente <651 THEN '535-650'
                                     WHEN ScoreSinacofiCliente >=651 AND ScoreSinacofiCliente <701 THEN '651-700'
                                     WHEN ScoreSinacofiCliente >=701 AND ScoreSinacofiCliente <751 THEN '701-750'
                                     WHEN ScoreSinacofiCliente >=751                                            THEN '751-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE       C_SIS_LABEL='CF' AND    N_LOGICPATH_LABEL IN ('Nuevo Campana')-- AND CATEG_SINACOFI IS NULL ORDER BY          ScoreSinacofiCliente      

            -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI

UPDATE #CALCULO_RI 
 SET CATEG_SCORE_INT= CASE WHEN SCORE_CF_ORI_CALCULADO <541                                          THEN '000-540'
                                     WHEN SCORE_CF_ORI_CALCULADO >=541 AND SCORE_CF_ORI_CALCULADO <551 THEN '541-550'
                                     WHEN SCORE_CF_ORI_CALCULADO >=551 AND SCORE_CF_ORI_CALCULADO <567 THEN '551-566'
                                     WHEN SCORE_CF_ORI_CALCULADO >=567 AND SCORE_CF_ORI_CALCULADO <581 THEN '567-580'
                                     WHEN SCORE_CF_ORI_CALCULADO >=581 AND SCORE_CF_ORI_CALCULADO <591 THEN '581-590'
                                     WHEN SCORE_CF_ORI_CALCULADO >=591 AND SCORE_CF_ORI_CALCULADO <611 THEN '591-610'
                                     WHEN SCORE_CF_ORI_CALCULADO >=611 AND SCORE_CF_ORI_CALCULADO <621 THEN '611-620'
                                     WHEN SCORE_CF_ORI_CALCULADO >=621 AND SCORE_CF_ORI_CALCULADO <631 THEN '621-630'
                                     WHEN SCORE_CF_ORI_CALCULADO >=631                              THEN '631-999' END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
WHERE       C_SIS_LABEL='CF' AND    N_LOGICPATH_LABEL IN ('Nuevo Campana')
--AND CATEG_SINACOFI IS NULL 
 --ORDER BY SCORE_PB_ORI_CALCULADO   
 

  
   -- SELECT * FROM RT_SCORING.DBO.CALCULO_RI

UPDATE #CALCULO_RI 
 SET RI= CASE                WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='541-550'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='551-566'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='567-580'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='581-590'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='591-610'  THEN 'E'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='611-620'  THEN 'C'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='621-630'  THEN 'C'
                                   WHEN CATEG_SINACOFI='000-347'  AND CATEG_SCORE_INT='631-999'  THEN 'B'
                                                                                                          
                                    WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='541-550'  THEN 'E'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='551-566'  THEN 'E'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='567-580'  THEN 'E'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='581-590'  THEN 'D'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='591-610'  THEN 'D'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='611-620'  THEN 'B'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='621-630'  THEN 'B'
                                   WHEN CATEG_SINACOFI='348-400' AND CATEG_SCORE_INT='631-999'  THEN 'B'
                                                                                                          
                                    WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='541-550'  THEN 'E'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='551-566'  THEN 'E'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='567-580'  THEN 'D'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='581-590'  THEN 'D'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='591-610'  THEN 'C'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='611-620'  THEN 'B'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='621-630'  THEN 'B'
                                   WHEN CATEG_SINACOFI='401-534' AND CATEG_SCORE_INT='631-999'  THEN 'B'
                                                                                                          
                                    WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='541-550'  THEN 'E'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='551-566'  THEN 'E'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='567-580'  THEN 'D'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='581-590'  THEN 'D'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='591-610'  THEN 'C'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='611-620'  THEN 'B'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='621-630'  THEN 'B'
                                   WHEN CATEG_SINACOFI='535-650'  AND CATEG_SCORE_INT='631-999'  THEN 'A'
                                                                                                          
                                    WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='541-550'  THEN 'D'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='551-566'  THEN 'D'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='567-580'  THEN 'C'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='581-590'  THEN 'C'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='591-610'  THEN 'B'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='611-620'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='621-630'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '651-700' AND CATEG_SCORE_INT='631-999'  THEN 'A'
                                   
                                    WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='541-550'  THEN 'D'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='551-566'  THEN 'C'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='567-580'  THEN 'C'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='581-590'  THEN 'B'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='591-610'  THEN 'B'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='611-620'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='621-630'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '701-750' AND CATEG_SCORE_INT='631-999'  THEN 'A'
                                   
                                    WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='000-540'  THEN 'E'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='541-550'  THEN 'D'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='551-566'  THEN 'B'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='567-580'  THEN 'B'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='581-590'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='591-610'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='611-620'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='621-630'  THEN 'A'
                                   WHEN CATEG_SINACOFI= '751-999' AND CATEG_SCORE_INT='631-999'  THEN 'A'
                                                                      END-- SELECT * FROM RT_SCORING.DBO.CALCULO_RI 
                                   WHERE       C_SIS_LABEL='CF' AND    N_LOGICPATH_LABEL IN ('Nuevo Campana')       
