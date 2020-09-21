import cx_Oracle
import sys


def sysdate():
    con = cx_Oracle.connect('SPUSER', 'SPUSER', '210.216.37.5:1521/SPRING')

    cur = con.cursor()
    '''
    sql = "SELECT TO_CHAR(SYSDATE - 7, 'YYYYMMDDHH24MISS')," \
          "TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS')," \
          "TO_CHAR(TRUNC(TO_DATE(TO_DATE(TO_CHAR(SYSDATE,'YYYY')||'01', 'yyyymm')-7), 'IW') + 4, 'yyyymmdd')||'060000'," \
          "TO_CHAR(SYSDATE-7, 'IW') WW " \
          "FROM DUAL"
    '''
    '''
    sql = "SELECT TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS')," \
          "TO_CHAR(SYSDATE-3, 'YYYYMMDDHH24MISS')," \
          "TO_CHAR(TRUNC(TO_DATE(TO_DATE(TO_CHAR(SYSDATE,'YYYY')||'01', 'yyyymm')), 'IW') + 4, 'yyyymmdd')||'060000'," \
          "TO_CHAR(SYSDATE-3, 'IW') WW " \
          "FROM DUAL"
    '''

    sql = "select '20200911080000', '20200918080000' from dual"
    cur.execute(sql)

    date = cur.fetchall()

    cur.close()

    con.close()

    return date


def item_query(ST_dt,END_dt):
    con = cx_Oracle.connect('SPUSER', 'SPUSER', '210.216.37.5:1521/SPRING')

    cur = con.cursor()

    # Select 문장
    sql = """
            SELECT * FROM(
SELECT PRA_CD, CMP_NAME, FAMILY, PKG, PKG_OPT, PART_NO, SALE_CD, FNP_GETDIECD(ITEM_CD) DIE_CODE, FNP_GETRELNO(RUN_NO) REL_NO, RUN_NO, FNP_GETWMLOTNO(LOT_NO)  SPLIT_RUN_NO,
         LOT_NO, MAT_YN, MCN_CD, IN_QTY, OUT_QTY, MV_START_DT, MV_END_DT,99.1, FNQ_GETLOSSENAME(LS_CD) LS_NM, LS_QTY
    FROM (     SELECT A.PRA_CD       PRA_CD,
                      E.CMP_SNM      CMP_NAME,
                 G.ITEM_LNM     FAMILY,
                      D.ITEM_MNM     PKG,
                      FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,
                      C.PART_NO      PART_NO,
                      C.SALE_CD      SALE_CD,
                      F.WM_RUN_NO    RUN_NO,
                      A.LOT_NO       LOT_NO,
                      A.MV_IN_QTY - NVL(FNP_GETCUSTLSTOTAL(A.LOT_NO,A.PRA_CD),0) IN_QTY,
                      A.MV_OUT_QTY   OUT_QTY,
                      A.MCN_CD       MCN_CD,
                      /*B.LS_CD        LS_CD,*/
                      CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN '' ELSE B.LS_CD END LS_CD,
                      /*SUM(NVL(B.LS_QTY,0)) LS_QTY,*/
                 CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN 0 ELSE SUM(NVL(B.LS_QTY,0)) END LS_QTY,
                      NVL(F.WM_MAT_YN,'N') MAT_YN,
                 A.ITEM_CD      ITEM_CD,
                      (SELECT DISTINCT SUBSTR(Q.MV_START_DT, 1, 4) || '-' || SUBSTR(Q.MV_START_DT, 5, 2) || '-' || SUBSTR(Q.MV_START_DT, 7, 2) || ' ' || SUBSTR(Q.MV_START_DT, 9, 2) || ':' || SUBSTR(Q.MV_START_DT, 11, 2) || ':' || SUBSTR(Q.MV_START_DT, 13, 2)
                        FROM TP_MV Q WHERE Q.PRA_CD = 'A020' AND Q.LOT_NO = A.LOT_NO) MV_START_DT,
                 SUBSTR(A.MV_END_DT, 1, 4) || '-' || SUBSTR(A.MV_END_DT, 5, 2) || '-' || SUBSTR(A.MV_END_DT, 7, 2) || ' ' || SUBSTR(A.MV_END_DT, 9, 2) || ':' || SUBSTR(A.MV_END_DT, 11, 2) || ':' || SUBSTR(A.MV_END_DT, 13, 2) MV_END_DT
               FROM TP_MV A, TP_LS B, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TP_WM F, TD_ITEMLGRP G
                WHERE  ( A.MV_END_DT >= '%s' )
                    AND ( A.MV_END_DT <= '%s' )
                    AND ( A.PRA_CD    ='T010')
                AND ( DECODE(A.MV_OUT_QTY, 0 , 1, A.MV_OUT_QTY) / DECODE(NVL(A.MV_IN_QTY, 0), 0, 1, A.MV_IN_QTY)) * 100  >= ('0.0')
                    AND ( A.BSN_CD       = B.BSN_CD(+) )
                    AND ( A.LOT_NO       = B.LOT_NO(+) )
                    AND ( A.PRA_CD       = B.PRA_CD(+) )
                    AND ( A.MV_TRANS_CNT = B.LS_TRANS_CNT(+) )
                    AND ( A.ITEM_CD      = C.ITEM_CD )
                    AND ( C.CMP_CD    ='002011')
                    AND ( C.ITEM_MCD     = D.ITEM_MCD )
               AND ( C.ITEM_LCD     = G.ITEM_LCD )
                    AND ( C.CMP_CD       = E.CMP_CD )
                    AND ( A.BSN_CD       = F.BSN_CD )
                    AND ( A.LOT_NO       = F.LOT_NO )
                    AND ( F.WM_TRANS_CNT = (SELECT MAX(WM_TRANS_CNT) FROM TP_WM WHERE BSN_CD = A.BSN_CD AND  LOT_NO = A.LOT_NO) )
                GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD, A.ITEM_CD, F.WM_RUN_NO,
                     A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY, B.LS_CD, A.MCN_CD ,F.WM_MAT_YN, MV_START_DT, MV_END_DT
                UNION ALL
                SELECT 'SPOS' PRA_CD,
                        FNP_GETCMPNM(X.CMP_CD) CMP_NAME,
                        X.ITEM_LNM FAMILY,
                        FNP_GETPACKAGE(X.ITEM_CD) PKG,
                        FNP_GETCDNM(FNP_GETPKGOPT(X.ITEM_CD)) PKG_OPT,
                        X.PART_NO PART_NO,
                        X.SALE_CD SALE_CD,
                        FNP_GETCUSTLOTNO(X.LOT_NO) RUN_NO,
                        X.LOT_NO LOT_NO,
                        X.IN_QTY IN_QTY,
                        X.OUT_QTY OUT_QTY,
                        '' MCN_CD,
                        Y.LS_CD LS_CD,
                        SUM(Y.LS_QTY) LS_QTY,
                        FNP_GETMATYN(X.LOT_NO) MAT_YN,
                        X.ITEM_CD,
                        '' MV_START_DT,
                       '' MV_END_DT
                 FROM(
                        SELECT  /*+ INDEX_DESC(A TP_WT_IDX02) */
                                'SPOS' PRA_CD, A.CMP_CD, G.ITEM_LNM, A.ITEM_CD, C.PART_NO, C.SALE_CD, A.LOT_NO, A.WT_CURR_QTY + NVL(FNP_GETLSQTY(A.LOT_NO, 'SPOS'), 0) IN_QTY, A.WT_CURR_QTY OUT_QTY
                        FROM TP_WT A, TD_ITEMMAS C, TD_ITEMLGRP G
                         WHERE (A.PRA_CD = 'A050')
                            AND (A.WT_STATUS = 'PFOO')
                            AND (A.WT_ACT_DT >= '%s')
                            AND (A.WT_ACT_DT <  '%s')
                            AND (C.ITEM_LCD = G.ITEM_LCD)
                            AND ( C.ITEM_MCD  = '*')
                            AND ( C.ITEM_TYPE = '*')
                            AND ( A.BSN_CD    ='S1')
                            )X, TP_LS Y
                 WHERE ( X.LOT_NO       = Y.LOT_NO(+) )
                    AND ( X.PRA_CD       = Y.PRA_CD(+) )
                    AND ( DECODE(X.OUT_QTY, 0 , 1, X.OUT_QTY) / DECODE(NVL(X.IN_QTY, 0), 0, 1, X.IN_QTY)) * 100  >= ('0.0')
                    AND ( X.CMP_CD    ='002011')
                    AND ( X.SALE_CD   = '*')
                    AND NOT X.PRA_CD IN (DECODE('N' , 'N', 'SPOS', ' '))
                 GROUP BY FNP_GETCMPNM(X.CMP_CD), X.ITEM_LNM, FNP_GETPACKAGE(X.ITEM_CD),FNP_GETCDNM(FNP_GETPKGOPT(X.ITEM_CD)),X.PART_NO,X.SALE_CD,
                          FNP_GETCUSTLOTNO(X.LOT_NO), X.LOT_NO, X.IN_QTY, X.OUT_QTY, Y.LS_CD, FNP_GETMATYN(X.LOT_NO), X.ITEM_CD
             UNION ALL
                SELECT A.PRA_CD       PRA_CD,
                      E.CMP_SNM      CMP_NAME,
                 G.ITEM_LNM     FAMILY,
                      D.ITEM_MNM     PKG,
                      FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,
                      C.PART_NO      PART_NO,
                      C.SALE_CD      SALE_CD,
                      F.WM_RUN_NO    RUN_NO,
                      A.LOT_NO       LOT_NO,
                      A.WT_PRE_QTY - NVL(FNP_GETCUSTLSTOTAL(A.LOT_NO,A.PRA_CD),0) IN_QTY,
                      A.WT_CURR_QTY   OUT_QTY,
                      '*'       MCN_CD,
                      /*B.LS_CD        LS_CD,*/
                      CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN '' ELSE B.LS_CD END LS_CD,
                      /*SUM(NVL(B.LS_QTY,0)) LS_QTY ,*/
                 CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN 0 ELSE SUM(NVL(B.LS_QTY,0)) END LS_QTY,
                        NVL(F.WM_MAT_YN,'N')  MAT_YN  ,
                 A.ITEM_CD      ITEM_CD,
                     '' MV_START_DT,
                      '' MV_END_DT
               FROM TP_WT A, TP_LS B, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TP_WM F, TD_ITEMLGRP G
                WHERE  ( A.WT_ACT_DT >= '%s' )
                    AND ( A.WT_ACT_DT <= '%s' )
                    AND ( A.PRA_CD    ='T010')
                AND ( DECODE(A.WT_CURR_QTY, 0 , 0, A.WT_CURR_QTY) / DECODE(NVL(A.WT_PRE_QTY, 0), 0, 0, A.WT_PRE_QTY)) * 100  >= ('0.0')
                    AND ( A.BSN_CD       = B.BSN_CD(+) )
                    AND ( A.LOT_NO       = B.LOT_NO(+) )
                    AND ( A.PRA_CD       = B.PRA_CD(+) )
                    AND ( A.WT_TRANS_CNT = B.LS_TRANS_CNT(+) )
                    AND ( A.ITEM_CD      = C.ITEM_CD )
                    AND ( C.CMP_CD    ='002011')
                    AND ( C.ITEM_MCD  = '*')
                    AND ( C.SALE_CD   = '*')
                    AND ( C.ITEM_TYPE = '*')
                    AND ( C.ITEM_MCD     = D.ITEM_MCD )
               AND ( C.ITEM_LCD     = G.ITEM_LCD )
                    AND ( C.CMP_CD       = E.CMP_CD )
                    AND ( A.BSN_CD       = F.BSN_CD )
                    AND ( A.LOT_NO       = F.LOT_NO )
                    AND ( F.WM_TRANS_CNT = (SELECT MAX(WM_TRANS_CNT) FROM TP_WM WHERE BSN_CD = A.BSN_CD AND  LOT_NO = A.LOT_NO) )
                    AND A.WT_CURR_QTY = 0
                     AND WT_STATUS = 'PFLS'
                GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD, A.ITEM_CD, F.WM_RUN_NO,
                     A.LOT_NO, A.WT_PRE_QTY, A.WT_CURR_QTY, B.LS_CD, '*' ,F.WM_MAT_YN
            UNION ALL
            SELECT A.PRA_CD             PRA_CD,
                     '***'                CMP_NAME,
                     '***'                FAMILY,
                     '***'                PKG,
                     '***'                     PKG_OPT,
                     '***'                PART_NO,
                     '***'                SALE_CD,
                     F.WM_RUN_NO          RUN_NO,
                     A.LOT_NO             LOT_NO,
                     A.MV_IN_QTY - NVL(FNP_GETCUSTLSTOTAL(A.LOT_NO,A.PRA_CD),0) IN_QTY,
                     A.MV_OUT_QTY         OUT_QTY,
                     A.MCN_CD             MCN_CD,
                     /*B.LS_CD              LS_CD,*/
                     CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN '' ELSE B.LS_CD END LS_CD,
                     /*SUM(NVL(B.LS_QTY,0)) LS_QTY,*/
                     CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN 0 ELSE SUM(NVL(B.LS_QTY,0)) END LS_QTY,
                      NVL(F.WM_MAT_YN,'N') MAT_YN,
                 '***'               ITEM_CD,
                      '' MV_START_DT,
                      '' MV_END_DT
                 FROM TP_MV A, TP_LS B, TP_WM F
                WHERE  ( A.MV_END_DT >= '%s' )
                    AND ( A.MV_END_DT <= '%s'   )
                    AND ( A.PRA_CD    ='T010')
                AND ( DECODE(A.MV_OUT_QTY, 0 , 1, A.MV_OUT_QTY) / DECODE(NVL(A.MV_IN_QTY, 0), 0, 1, A.MV_IN_QTY)) * 100  >= ('0.0')
                    AND ( A.BSN_CD       = B.BSN_CD(+) )
                    AND ( A.LOT_NO       = B.LOT_NO(+) )
                    AND ( A.PRA_CD       = B.PRA_CD(+) )
                    AND ( A.MV_TRANS_CNT = B.LS_TRANS_CNT(+) )
                    AND ( A.ITEM_CD      = '*'      )
                    AND ( A.CMP_CD    ='002011')
                    AND ( A.BSN_CD    = F.BSN_CD   )
                    AND ( A.LOT_NO    = F.LOT_NO   )
                    AND ( F.WM_TRANS_CNT = (SELECT MAX(WM_TRANS_CNT) FROM TP_WM WHERE BSN_CD = A.BSN_CD AND LOT_NO = A.LOT_NO))
                GROUP BY A.PRA_CD, F.WM_RUN_NO, A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY, B.LS_CD, A.MCN_CD ,F.WM_MAT_YN)
                WHERE  PRA_CD IN DECODE('N' , 'Y', 'SPOS', PRA_CD)
    GROUP BY PRA_CD, CMP_NAME, FAMILY, PKG, PKG_OPT, PART_NO, SALE_CD, FNP_GETDIECD(ITEM_CD), RUN_NO, FNP_GETRELNO(RUN_NO), LOT_NO, IN_QTY, OUT_QTY, MCN_CD, LS_CD,MAT_YN, MV_START_DT, MV_END_DT, LS_QTY
    )    
"""%(ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt)

    cur.execute(sql)
    result = cur.fetchall()

    cur.close()

    con.close()

    return result


def detail_row(ST_dt,END_dt):

    con = cx_Oracle.connect('SPUSER', 'SPUSER', '210.216.37.5:1521/SPRING')

    cur = con.cursor()

    # Select 문장
    sql = """
        SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                                 
       FNS_SALECODE(ITEM_CD) SALE_CD,                             
       FNP_GETPACKAGE(ITEM_CD) PKG,                               
       FNP_GETDIECD(ITEM_CD) DIE_CD,                              
       FNP_GETFAMILY(ITEM_CD) FAMILY,                             
      '01_Reception date' PRA_CD,                             
       TO_NUMBER(MIN(SUBSTR(SP_aCT_DT,1,8))) QTY                              
FROM TP_SP 
WHERE LOT_NO IN(                              
                    SELECT FNP_GETCUSTLOTNO(LOT_NO) RUN_NO                                   
                    FROM TP_MV                                   
                    WHERE MV_END_DT >= '%s'                                  
                    AND MV_END_DT <= '%s'                                
                    AND PRA_CD ='T010'                                   
                    AND CMP_CD = '002011'
                    )                                
AND PRA_CD = 'A010'                                  
GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), FNS_SALECODE(ITEM_CD),  FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD) , FNP_GETFAMILY(ITEM_CD)
UNION ALL                                
SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                                
                    FNS_SALECODE(ITEM_CD) SALE_CD,                            
                    FNP_GETPACKAGE(ITEM_CD) PKG,                              
                    FNP_GETDIECD(ITEM_CD) DIE_CD,                             
                    FNP_GETFAMILY(ITEM_CD) FAMILY,                            
                    '02_Shipment date' PRA_CD,                            
                    TO_NUMBER(MAX(SUBSTR(SH_ACT_DT,1,8))) QTY                             
FROM TP_SH                               
WHERE SH_ACT_DT >= '%s'                                  
AND SH_ACT_DT <= '%s'                             
AND CMP_cD = '002011'                                
AND FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) IN(                               
                                                SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) 
                                                FROM TP_MV 
                                                WHERE MV_END_DT >= '%s' 
                                                AND MV_END_DT <= '%s' 
                                                AND PRA_CD ='T010' 
                                                AND CMP_CD = '002011'
                                                )
GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), FNS_SALECODE(ITEM_CD),FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD), FNP_GETFAMILY(ITEM_CD)
UNION ALL                                
SELECT  A.REL_NO,                                
        SALE_CD,                              
        PKG,                              
        DIE_CD,                               
        FAMILY,                               
        '03_INPUT_QTY' PRA_CD,                                
        A.OUT_QTY + NVL(B.LS_QTY,0) QTY                               
FROM(
        SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                                
        FNS_SALECODE(ITEM_CD) SALE_CD,                      
        FNP_GETPACKAGE(ITEM_CD) PKG,                        
        FNP_GETDIECD(ITEM_CD) DIE_CD,                       
        FNP_GETFAMILY(ITEM_CD) FAMILY,                      
        PRA_CD,                         
        SUM(MV_OUT_QTY) OUT_QTY                         
        FROM TP_MV                         
        WHERE LOT_NO IN(
                            SELECT LOT_NO                          
                            FROM TP_MV                 
                            WHERE MV_END_DT >= '%s'                
                            AND MV_END_DT <= '%s'              
                            AND PRA_CD ='T010'                 
                            AND CMP_CD = '002011'
                            )              
        AND PRA_CD = 'A020'                            
        GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), PRA_CD,FNS_SALECODE(ITEM_CD), FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD), FNP_GETFAMILY(ITEM_CD)) A,                          
        (   
            SELECT  FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                            
                    PRA_CD,                         
                    NVL(SUM(LS_QTY),0) LS_QTY                       
            FROM TP_LS                             
            WHERE LOT_NO IN(    
                                SELECT LOT_NO                          
                                FROM TP_MV                 
                                WHERE MV_END_DT >= '%s'                
                                AND MV_END_DT <= '%s'              
                                AND PRA_CD ='T010'                 
                                AND CMP_CD = '002011'
                                )              
            AND PRA_CD = 'A020'                
            AND LS_CD = 'BE055'                
            GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), PRA_CD
            ) B                
WHERE A.REL_NO = B.REL_NO(+)                                 
AND A.PRA_CD = B.PRA_CD(+)                               
UNION ALL                                
SELECT  REL_NO,                                  
        SALE_CD,                              
        PKG,                              
        DIE_CD,                               
        FAMILY,                               
        '04_BALANCE_QTY' PRA_CD,                              
        LS_QTY QTY                            
FROM(   
        SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                                
        PRA_CD,                         
        FNS_SALECODE(ITEM_CD) SALE_CD,                      
        FNP_GETPACKAGE(ITEM_CD) PKG,                        
        FNP_GETDIECD(ITEM_CD) DIE_CD,                       
        FNP_GETFAMILY(ITEM_CD) FAMILY,                      
        NVL(SUM(LS_QTY),0) LS_QTY                       
        FROM TP_LS                                
        WHERE LOT_NO IN(
                            SELECT LOT_NO                          
                            FROM TP_MV                 
                            WHERE MV_END_DT >= '%s'                
                            AND MV_END_DT <= '%s'              
                            AND PRA_CD ='T010'                 
                            AND CMP_CD = '002011'
                            )              
        AND PRA_CD = 'A020'                
        AND LS_CD = 'BE055'                
        GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), PRA_CD, FNS_SALECODE(ITEM_CD),  FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD), FNP_GETFAMILY(ITEM_CD)
        )              
UNION ALL                                
SELECT  A.REL_NO,                                
        A.SALE_CD,                            
        A.PKG,                            
        A.DIE_CD,                             
        A.FAMILY,                             
        A.PRA_CD,                             
        A.IN_QTY - A.OUT_QTY - NVL(B.LS_QTY,0) QTY                                
FROM(
        SELECT  FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                               
                PRA_CD,                         
                FNS_SALECODE(ITEM_CD) SALE_CD,                      
                FNP_GETPACKAGE(ITEM_CD) PKG,                        
                FNP_GETDIECD(ITEM_CD) DIE_CD,                       
                FNP_GETFAMILY(ITEM_CD) FAMILY,                      
                SUM(MV_IN_QTY) IN_QTY,                          
                SUM(MV_OUT_QTY) OUT_QTY                         
        FROM TP_MV                         
        WHERE LOT_NO IN(
                            SELECT LOT_NO                          
                            FROM TP_MV                 
                            WHERE MV_END_DT >= '%s'                
                            AND MV_END_DT <= '%s'              
                            AND PRA_CD ='T010'                 
                            AND CMP_CD = '002011'
                            )              
        AND PRA_CD ='A020'                         
        GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), PRA_CD,FNS_SALECODE(ITEM_CD), FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD), FNP_GETFAMILY(ITEM_CD)) A,                          
        (
            SELECT  FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                            
                    PRA_CD,                         
                    NVL(SUM(LS_QTY),0) LS_QTY                       
            FROM TP_LS                             
            WHERE LOT_NO IN(
                                SELECT LOT_NO                          
                                FROM TP_MV                 
                                WHERE MV_END_DT >= '%s'                
                                AND MV_END_DT <= '%s'              
                                AND PRA_CD ='T010'                 
                                AND CMP_CD = '002011'
                                )              
            AND PRA_CD = 'A020'                
            AND LS_CD = 'BE055'                
            GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), PRA_CD
            ) B                
WHERE A.REL_NO = B.REL_NO(+)                                 
AND A.PRA_CD = B.PRA_CD(+)                               
UNION ALL                                
SELECT  REL_NO,                                  
        SALE_CD,                              
        PKG,                              
        DIE_CD,                               
        FAMILY,                               
        PRA_CD,                               
        IN_QTY - OUT_QTY LS_SUM                               
FROM(
        SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                                
        PRA_CD,                         
        FNS_SALECODE(ITEM_CD) SALE_CD,                      
        FNP_GETPACKAGE(ITEM_CD) PKG,                        
        FNP_GETDIECD(ITEM_CD) DIE_CD,                       
        FNP_GETFAMILY(ITEM_CD) FAMILY,                      
        SUM(MV_IN_QTY) IN_QTY,                          
        SUM(MV_OUT_QTY) OUT_QTY                         
        FROM TP_MV                         
        WHERE LOT_NO IN(
                            SELECT LOT_NO                          
                            FROM TP_MV                 
                            WHERE MV_END_DT >= '%s'                
                            AND MV_END_DT <= '%s'              
                            AND PRA_CD ='T010'                 
                            AND CMP_CD = '002011'
                            )              
        AND PRA_CD <>'A020'                            
        GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), PRA_CD, FNS_SALECODE(ITEM_CD), FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD), FNP_GETFAMILY(ITEM_CD)
        )                             
UNION ALL                                
SELECT  FNP_GETRELNO(FNP_GETCUSTLOTNO(A.LOT_NO)) REL_NO,                                 
        FNS_SALECODE(A.ITEM_CD) SALE_CD,                              
        FNP_GETPACKAGE(A.ITEM_CD) PKG,                            
        FNP_GETDIECD(A.ITEM_CD) DIE_CD,                               
        FNP_GETFAMILY(A.ITEM_CD) FAMILY,                              
        A.PRA_CD,                             
        SUM(A.WT_PRE_QTY) - SUM(A.WT_CURR_QTY) IN_QTY                             
FROM TP_WT A, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TD_ITEMLGRP G                               
WHERE A.LOT_NO IN (
                    SELECT LOT_NO                                
                    FROM TP_MV                                
                    WHERE MV_END_DT >= '%s'                   
                    AND MV_END_DT <= '%s'                 
                    AND PRA_CD ='T010'                    
                    AND CMP_CD = '002011'
                    )                 
AND (A.ITEM_CD      = C.ITEM_CD )                                
AND (A.CMP_CD = '002011')                                
AND (C.ITEM_MCD     = D.ITEM_MCD )                               
AND (C.ITEM_LCD     = G.ITEM_LCD )                               
AND (C.CMP_CD       = E.CMP_CD )                                 
AND A.WT_CURR_QTY = 0                                
AND WT_STATUS = 'PFLS'                               
GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD, A.LOT_NO, A.WT_PRE_QTY, A.WT_CURR_QTY,
FNS_SALECODE(A.ITEM_CD), FNP_GETPACKAGE(A.ITEM_CD), FNP_GETDIECD(A.ITEM_CD), FNP_GETFAMILY(A.ITEM_CD)                                
UNION ALL                                
SELECT  FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) REL_NO,                               
        FNS_SALECODE(ITEM_CD) SALE_CD,                            
        FNP_GETPACKAGE(ITEM_CD) PKG,                              
        FNP_GETDIECD(ITEM_CD) DIE_CD,                             
        FNP_GETFAMILY(ITEM_CD) FAMILY,                            
        'Z_Shipped Qty' PRA_CD,                               
        SUM(SH_QTY) QTY                               
FROM TP_SH                               
WHERE SH_ACT_DT >= '%s'                                  
AND SH_ACT_DT <= '%s'                             
AND CMP_cD = '002011'                                
AND FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) IN(
                                                SELECT FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)) 
                                                FROM TP_MV 
                                                WHERE MV_END_DT >= '%s' 
                                                AND MV_END_DT <= '%s' 
                                                AND PRA_CD ='T010' 
                                                AND CMP_CD = '002011'
                                                )
GROUP BY FNP_GETRELNO(FNP_GETCUSTLOTNO(LOT_NO)), FNS_SALECODE(ITEM_CD),FNP_GETPACKAGE(ITEM_CD), FNP_GETDIECD(ITEM_CD), FNP_GETFAMILY(ITEM_CD)

"""%(ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt)

    cur.execute(sql)
    result = cur.fetchall()

    # for result in cur:
    #    print(result)

    cur.close()

    con.close()


    return result



def rpt_l_row(rpt_l_ST_dt,rpt_l_END_dt):
    con = cx_Oracle.connect('SPUSER', 'SPUSER', '210.216.37.5:1521/SPRING')

    cur = con.cursor()

    # Select 문장
    sql = """
            SELECT  SH_GUBUN,
                    A.CMP_CD,
                    FNP_GETPACKAGE(A.ITEM_CD) PKG,
                    NVL(FNP_GETPKGOPT(A.ITEM_CD), '*') PKGOPT,
                    A.SALE_CD,
                    FNP_GETCUSTLOTNO(SP_LOTNO) RUN_NO,
                    CASE WHEN FNP_GETRCHCHIP(FNP_GETCUSTLOTNO(SP_LOTNO)) = '*' THEN 0 ELSE TO_NUMBER(FNP_GETRCHCHIP(FNP_GETCUSTLOTNO(SP_LOTNO))) END RUN_QTY,
                    SP_LOTNO,
                    LOT_NO,
                    INSP_QTY,
                    LS_QTY,
                    LS_CD,
                    BAD_TYPE,
                    BAD_NO,
                    MMCL_SEC,
                    REGULAR_YN,
                    JUDGE,
                    MCN_CD,
                    INSP_DT,
                    FNA_EMPNM(INSP_EMP) EMP_NM,
                    WKG_CD,
                    WEEK_CD,
                    B.ITEM_TYPE,
                    UPDATE_EMP,
                    UPDATE_DT,
                    BIGO       
            FROM TQ_SHINSP A, TD_ITEMMAS B
            WHERE A.ITEM_CD = B.ITEM_CD
            AND BSN_CD = 'S1'
            AND INSP_DT >= '%s'
            AND INSP_DT < '%s'           
            AND (A.CMP_CD = '002011')
        """%(rpt_l_ST_dt,rpt_l_END_dt)

    cur.execute(sql)
    result = cur.fetchall()

    # for result in cur:
    #    print(result)

    cur.close()

    con.close()

    return result

# IXYS(미국)-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
def item_query_P(ST_dt,END_dt):

    con = cx_Oracle.connect('SPUSER', 'SPUSER', '210.216.37.5:1521/SPRING')

    cur = con.cursor()

    # Select 문장
    sql = """    
       SELECT * FROM(
SELECT PKG, SALE_CD, FNP_GETWMLOTNO(LOT_NO)  SPLIT_RUN_NO, LOT_NO, IN_QTY, OUT_QTY, MV_END_DT, FNQ_GETLOSSENAME(LS_CD) LS_NM, LS_QTY
    FROM (     SELECT A.PRA_CD       PRA_CD,
                      E.CMP_SNM      CMP_NAME,
                      G.ITEM_LNM     FAMILY,
                      D.ITEM_MNM     PKG,
                      FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,
                      C.PART_NO      PART_NO,
                      C.SALE_CD      SALE_CD,
                      F.WM_RUN_NO    RUN_NO,
                      A.LOT_NO       LOT_NO,
                      A.MV_IN_QTY - NVL(FNP_GETCUSTLSTOTAL(A.LOT_NO,A.PRA_CD),0) IN_QTY,
                      A.MV_OUT_QTY   OUT_QTY,
                      A.MCN_CD       MCN_CD,
                      /*B.LS_CD        LS_CD,*/
                      CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN '' ELSE B.LS_CD END LS_CD,
                      /*SUM(NVL(B.LS_QTY,0)) LS_QTY,*/
                 CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN 0 ELSE SUM(NVL(B.LS_QTY,0)) END LS_QTY,
                      NVL(F.WM_MAT_YN,'N') MAT_YN,
                 A.ITEM_CD      ITEM_CD,
                      (SELECT DISTINCT SUBSTR(Q.MV_START_DT, 1, 4) || '-' || SUBSTR(Q.MV_START_DT, 5, 2) || '-' || SUBSTR(Q.MV_START_DT, 7, 2) || ' ' || SUBSTR(Q.MV_START_DT, 9, 2) || ':' || SUBSTR(Q.MV_START_DT, 11, 2) || ':' || SUBSTR(Q.MV_START_DT, 13, 2)
                        FROM TP_MV Q WHERE Q.PRA_CD = 'A020' AND Q.LOT_NO = A.LOT_NO) MV_START_DT,
                 SUBSTR(A.MV_END_DT, 1, 4) || '-' || SUBSTR(A.MV_END_DT, 5, 2) || '-' || SUBSTR(A.MV_END_DT, 7, 2) || ' ' || SUBSTR(A.MV_END_DT, 9, 2) || ':' || SUBSTR(A.MV_END_DT, 11, 2) || ':' || SUBSTR(A.MV_END_DT, 13, 2) MV_END_DT
               FROM TP_MV A, TP_LS B, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TP_WM F, TD_ITEMLGRP G
                WHERE  ( A.MV_END_DT >= '%s' )
                    AND ( A.MV_END_DT <= '%s' )
                    AND ( A.PRA_CD    ='T010')
                AND ( DECODE(A.MV_OUT_QTY, 0 , 1, A.MV_OUT_QTY) / DECODE(NVL(A.MV_IN_QTY, 0), 0, 1, A.MV_IN_QTY)) * 100  >= ('0.0')
                    AND ( A.BSN_CD       = B.BSN_CD(+) )
                    AND ( A.LOT_NO       = B.LOT_NO(+) )
                    AND ( A.PRA_CD       = B.PRA_CD(+) )
                    AND ( A.MV_TRANS_CNT = B.LS_TRANS_CNT(+) )
                    AND ( A.ITEM_CD      = C.ITEM_CD )
                    AND ( C.CMP_CD    ='002009')
                    AND ( C.ITEM_MCD     = D.ITEM_MCD )
               AND ( C.ITEM_LCD     = G.ITEM_LCD )
                    AND ( C.CMP_CD       = E.CMP_CD )
                    AND ( A.BSN_CD       = F.BSN_CD )
                    AND ( A.LOT_NO       = F.LOT_NO )
                    AND ( F.WM_TRANS_CNT = (SELECT MAX(WM_TRANS_CNT) FROM TP_WM WHERE BSN_CD = A.BSN_CD AND  LOT_NO = A.LOT_NO) )
                GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD, A.ITEM_CD, F.WM_RUN_NO,
                     A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY, B.LS_CD, A.MCN_CD ,F.WM_MAT_YN, MV_START_DT, MV_END_DT
                UNION ALL
                SELECT 'SPOS' PRA_CD,
                        FNP_GETCMPNM(X.CMP_CD) CMP_NAME,
                        X.ITEM_LNM FAMILY,
                        FNP_GETPACKAGE(X.ITEM_CD) PKG,
                        FNP_GETCDNM(FNP_GETPKGOPT(X.ITEM_CD)) PKG_OPT,
                        X.PART_NO PART_NO,
                        X.SALE_CD SALE_CD,
                        FNP_GETCUSTLOTNO(X.LOT_NO) RUN_NO,
                        X.LOT_NO LOT_NO,
                        X.IN_QTY IN_QTY,
                        X.OUT_QTY OUT_QTY,
                        '' MCN_CD,
                        Y.LS_CD LS_CD,
                        SUM(Y.LS_QTY) LS_QTY,
                        FNP_GETMATYN(X.LOT_NO) MAT_YN,
                        X.ITEM_CD,
                        '' MV_START_DT,
                       '' MV_END_DT
                 FROM(
                        SELECT  /*+ INDEX_DESC(A TP_WT_IDX02) */
                                'SPOS' PRA_CD, A.CMP_CD, G.ITEM_LNM, A.ITEM_CD, C.PART_NO, C.SALE_CD, A.LOT_NO, A.WT_CURR_QTY + NVL(FNP_GETLSQTY(A.LOT_NO, 'SPOS'), 0) IN_QTY, A.WT_CURR_QTY OUT_QTY
                        FROM TP_WT A, TD_ITEMMAS C, TD_ITEMLGRP G
                         WHERE (A.PRA_CD = 'A050')
                            AND (A.WT_STATUS = 'PFOO')
                            AND (A.WT_ACT_DT >= '%s')
                            AND (A.WT_ACT_DT <  '%s')
                            AND (C.ITEM_LCD = G.ITEM_LCD)
                            AND ( C.ITEM_MCD  = '*')
                            AND ( C.ITEM_TYPE = '*')
                            AND ( A.BSN_CD    ='S1')
                            )X, TP_LS Y
                 WHERE ( X.LOT_NO       = Y.LOT_NO(+) )
                    AND ( X.PRA_CD       = Y.PRA_CD(+) )
                    AND ( DECODE(X.OUT_QTY, 0 , 1, X.OUT_QTY) / DECODE(NVL(X.IN_QTY, 0), 0, 1, X.IN_QTY)) * 100  >= ('0.0')
                    AND ( X.CMP_CD    ='002009')
                    AND ( X.SALE_CD   = '*')
                    AND NOT X.PRA_CD IN (DECODE('N' , 'N', 'SPOS', ' '))
                 GROUP BY FNP_GETCMPNM(X.CMP_CD), X.ITEM_LNM, FNP_GETPACKAGE(X.ITEM_CD),FNP_GETCDNM(FNP_GETPKGOPT(X.ITEM_CD)),X.PART_NO,X.SALE_CD,
                          FNP_GETCUSTLOTNO(X.LOT_NO), X.LOT_NO, X.IN_QTY, X.OUT_QTY, Y.LS_CD, FNP_GETMATYN(X.LOT_NO), X.ITEM_CD
             UNION ALL
                SELECT A.PRA_CD       PRA_CD,
                      E.CMP_SNM      CMP_NAME,
                 G.ITEM_LNM     FAMILY,
                      D.ITEM_MNM     PKG,
                      FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,
                      C.PART_NO      PART_NO,
                      C.SALE_CD      SALE_CD,
                      F.WM_RUN_NO    RUN_NO,
                      A.LOT_NO       LOT_NO,
                      A.WT_PRE_QTY - NVL(FNP_GETCUSTLSTOTAL(A.LOT_NO,A.PRA_CD),0) IN_QTY,
                      A.WT_CURR_QTY   OUT_QTY,
                      '*'       MCN_CD,
                      /*B.LS_CD        LS_CD,*/
                      CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN '' ELSE B.LS_CD END LS_CD,
                      /*SUM(NVL(B.LS_QTY,0)) LS_QTY ,*/
                 CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN 0 ELSE SUM(NVL(B.LS_QTY,0)) END LS_QTY,
                        NVL(F.WM_MAT_YN,'N')  MAT_YN  ,
                 A.ITEM_CD      ITEM_CD,
                     '' MV_START_DT,
                      '' MV_END_DT
               FROM TP_WT A, TP_LS B, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TP_WM F, TD_ITEMLGRP G
                WHERE  ( A.WT_ACT_DT >= '%s' )
                    AND ( A.WT_ACT_DT <= '%s' )
                    AND ( A.PRA_CD    ='T010')
                AND ( DECODE(A.WT_CURR_QTY, 0 , 0, A.WT_CURR_QTY) / DECODE(NVL(A.WT_PRE_QTY, 0), 0, 0, A.WT_PRE_QTY)) * 100  >= ('0.0')
                    AND ( A.BSN_CD       = B.BSN_CD(+) )
                    AND ( A.LOT_NO       = B.LOT_NO(+) )
                    AND ( A.PRA_CD       = B.PRA_CD(+) )
                    AND ( A.WT_TRANS_CNT = B.LS_TRANS_CNT(+) )
                    AND ( A.ITEM_CD      = C.ITEM_CD )
                    AND ( C.CMP_CD    ='002009')
                    AND ( C.ITEM_MCD  = '*')
                    AND ( C.SALE_CD   = '*')
                    AND ( C.ITEM_TYPE = '*')
                    AND ( C.ITEM_MCD     = D.ITEM_MCD )
               AND ( C.ITEM_LCD     = G.ITEM_LCD )
                    AND ( C.CMP_CD       = E.CMP_CD )
                    AND ( A.BSN_CD       = F.BSN_CD )
                    AND ( A.LOT_NO       = F.LOT_NO )
                    AND ( F.WM_TRANS_CNT = (SELECT MAX(WM_TRANS_CNT) FROM TP_WM WHERE BSN_CD = A.BSN_CD AND  LOT_NO = A.LOT_NO) )
                    AND A.WT_CURR_QTY = 0
                     AND WT_STATUS = 'PFLS'
                GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD, A.ITEM_CD, F.WM_RUN_NO,
                     A.LOT_NO, A.WT_PRE_QTY, A.WT_CURR_QTY, B.LS_CD, '*' ,F.WM_MAT_YN
            UNION ALL
            SELECT A.PRA_CD             PRA_CD,
                     '***'                CMP_NAME,
                     '***'                FAMILY,
                     '***'                PKG,
                     '***'                     PKG_OPT,
                     '***'                PART_NO,
                     '***'                SALE_CD,
                     F.WM_RUN_NO          RUN_NO,
                     A.LOT_NO             LOT_NO,
                     A.MV_IN_QTY - NVL(FNP_GETCUSTLSTOTAL(A.LOT_NO,A.PRA_CD),0) IN_QTY,
                     A.MV_OUT_QTY         OUT_QTY,
                     A.MCN_CD             MCN_CD,
                     /*B.LS_CD              LS_CD,*/
                     CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN '' ELSE B.LS_CD END LS_CD,
                     /*SUM(NVL(B.LS_QTY,0)) LS_QTY,*/
                     CASE WHEN B.LS_CD IN('AS185', 'AS186', 'AS122') THEN 0 ELSE SUM(NVL(B.LS_QTY,0)) END LS_QTY,
                      NVL(F.WM_MAT_YN,'N') MAT_YN,
                 '***'               ITEM_CD,
                      '' MV_START_DT,
                      '' MV_END_DT
                 FROM TP_MV A, TP_LS B, TP_WM F
                WHERE  ( A.MV_END_DT >= '%s' )
                    AND ( A.MV_END_DT <= '%s'   )
                    AND ( A.PRA_CD    ='T010')
                AND ( DECODE(A.MV_OUT_QTY, 0 , 1, A.MV_OUT_QTY) / DECODE(NVL(A.MV_IN_QTY, 0), 0, 1, A.MV_IN_QTY)) * 100  >= ('0.0')
                    AND ( A.BSN_CD       = B.BSN_CD(+) )
                    AND ( A.LOT_NO       = B.LOT_NO(+) )
                    AND ( A.PRA_CD       = B.PRA_CD(+) )
                    AND ( A.MV_TRANS_CNT = B.LS_TRANS_CNT(+) )
                    AND ( A.ITEM_CD      = '*'      )
                    AND ( A.CMP_CD    ='002009')
                    AND ( A.BSN_CD    = F.BSN_CD   )
                    AND ( A.LOT_NO    = F.LOT_NO   )
                    AND ( F.WM_TRANS_CNT = (SELECT MAX(WM_TRANS_CNT) FROM TP_WM WHERE BSN_CD = A.BSN_CD AND LOT_NO = A.LOT_NO))
                GROUP BY A.PRA_CD, F.WM_RUN_NO, A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY, B.LS_CD, A.MCN_CD ,F.WM_MAT_YN)
                WHERE  PRA_CD IN DECODE('N' , 'Y', 'SPOS', PRA_CD)
    GROUP BY PRA_CD, CMP_NAME, FAMILY, PKG, PKG_OPT, PART_NO, SALE_CD, FNP_GETDIECD(ITEM_CD), RUN_NO, FNP_GETRELNO(RUN_NO), LOT_NO, IN_QTY, OUT_QTY, MCN_CD, LS_CD,MAT_YN, MV_START_DT, MV_END_DT, LS_QTY
    )
   """%(ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt )
    cur.execute(sql)

    result = cur.fetchall()

    cur.close()

    con.close()

    return result

def item_query_T(ST_dt,END_dt):

    con = cx_Oracle.connect('SPUSER', 'SPUSER', '210.216.37.5:1521/SPRING')

    cur = con.cursor()

    # Select 문장
    sql = """    
               SELECT PRA_CD, CMP_NAME, FAMILY, PKG, PKG_OPT, PART_NO, SALE_CD, RUN_NO, DECODE(FNP_GETWMLOTNO(LOT_NO), '', RUN_NO, FNP_GETWMLOTNO(LOT_NO) ) SPLIT_RUN_NO,                                                    
             LOT_NO, IN_QTY, OUT_QTY FROM (                                                                      
             SELECT    A.PRA_CD                    PRA_CD,                                                    
                          FNP_GETCMPNM(A.CMP_CD)      CMP_NAME,                                                    
                     FNP_GETFAMILY(A.ITEM_CD)    FAMILY,                                                    
                          FNP_GETPACKAGE(A.ITEM_CD)   PKG,                                                    
                          FNP_GETCDNM(B.ITEM_PKG_OPT) PKG_OPT,                                                    
                          FNP_GETPARTNO(A.ITEM_CD)     PART_NO,                                                    
                          FNS_SALECODE(A.ITEM_CD)      SALE_CD,                                                      
                          FNP_GETCUSTLOTNO(A.LOT_NO)    RUN_NO,                                                    
                          A.LOT_NO       LOT_NO,                                                                 
                          A.MV_IN_QTY  IN_QTY,                                                         
                          A.MV_OUT_QTY   OUT_QTY                                                    
                   FROM TP_MV A, TD_ITEMMAS B                                                    
                    WHERE   A.LOT_NO IN     (SELECT LOT_NO FROM TP_MV WHERE MV_END_DT >= '%s' AND MV_END_DT <= '%s'  AND PRA_CD ='T010')                                                    
                        AND ( A.CMP_CD    = '002009')                                                    
                        AND (A.ITEM_CD    = B.ITEM_CD)                                                    
                    GROUP BY A.PRA_CD, FNP_GETCMPNM(A.CMP_CD), FNP_GETFAMILY(A.ITEM_CD), FNP_GETPACKAGE(A.ITEM_CD), FNP_GETCDNM(B.ITEM_PKG_OPT),                                                    
                                FNP_GETPARTNO(A.ITEM_CD), FNS_SALECODE(A.ITEM_CD), FNP_GETCUSTLOTNO(A.LOT_NO), A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY                                                                           
              UNION ALL                                                                        
                        SELECT       A.PRA_CD       PRA_CD,                                                    
                          E.CMP_SNM      CMP_NAME,                                                    
                           G.ITEM_LNM     FAMILY,                                                    
                          D.ITEM_MNM     PKG,                                                    
                          FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,                                                    
                          C.PART_NO      PART_NO,                                                    
                          C.SALE_CD      SALE_CD,                                                     
                          FNP_GETCUSTLOTNO(A.LOT_NO)    RUN_NO,                                                    
                          A.LOT_NO       LOT_NO,                                                                 
                          A.MV_IN_QTY  IN_QTY,                                                         
                          A.MV_OUT_QTY   OUT_QTY                                                    
                   FROM TP_MV A, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TD_ITEMLGRP G                                                    
                    WHERE  A.LOT_NO IN                                                    
                           (                                                    
                           SELECT DISTINCT WM_LOT_NO FROM TP_AS                                                     
                            WHERE LOT_NO IN (SELECT LOT_NO FROM TP_MV WHERE MV_END_DT >= '%s' AND MV_END_DT <= '%s'  AND PRA_CD ='T010')                                                    
                           )                                                    
                        AND ( A.ITEM_CD      = C.ITEM_CD )                                                    
                        AND ( A.CMP_CD    = '002009')                                                    
                        AND ( C.ITEM_MCD     = D.ITEM_MCD )                                                    
                   AND ( C.ITEM_LCD     = G.ITEM_LCD )                                                    
                        AND ( C.CMP_CD       = E.CMP_CD )                                                    
                    GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD,                                                      
                             A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY                                                                      
              UNION ALL                                                       
              SELECT       A.PRA_CD       PRA_CD,                                                    
                          E.CMP_SNM      CMP_NAME,                                                    
                           G.ITEM_LNM     FAMILY,                                                    
                          D.ITEM_MNM     PKG,                                                    
                          FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,                                                    
                          C.PART_NO      PART_NO,                                                    
                          C.SALE_CD      SALE_CD,                                                     
                          FNP_GETCUSTLOTNO(A.LOT_NO)    RUN_NO,                                                    
                          A.LOT_NO       LOT_NO,                                                                 
                          A.MV_IN_QTY  IN_QTY,                                                         
                          A.MV_OUT_QTY   OUT_QTY                                                    
                   FROM TP_MV A, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TD_ITEMLGRP G                                                    
                    WHERE   A.LOT_NO IN                                                         
                           (SELECT CB_LOT_NO FROM TP_CB                                                     
                           WHERE LOT_NO IN (SELECT LOT_NO FROM TP_MV WHERE MV_END_DT >= '%s' AND MV_END_DT <= '%s'  AND PRA_CD ='T010')                                                    
                           AND   PRA_CD ='T010'                                                    
                            AND   CB_ACT_DT >= '%s' AND CB_ACT_DT <= '%s'                                                     
                           )                                                    
                        AND ( A.ITEM_CD      = C.ITEM_CD )                                                    
                        AND ( C.CMP_CD    = '002009')                                                    
                        AND ( C.ITEM_MCD     = D.ITEM_MCD )                                                    
                   AND ( C.ITEM_LCD     = G.ITEM_LCD )                                                    
                        AND ( C.CMP_CD       = E.CMP_CD )                                                    
                    GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD,                                                     
                             A.LOT_NO, A.MV_IN_QTY, A.MV_OUT_QTY                                                                     
                    UNION ALL                                                                
                    SELECT A.PRA_CD       PRA_CD,                                                    
                          E.CMP_SNM      CMP_NAME,                                                    
                     G.ITEM_LNM     FAMILY,                                                    
                          D.ITEM_MNM     PKG,                                                    
                          FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,                                                    
                          C.PART_NO      PART_NO,                                                    
                          C.SALE_CD      SALE_CD,                                                     
                          FNP_GETCUSTLOTNO(A.LOT_NO)    RUN_NO,                                                    
                          A.LOT_NO       LOT_NO,                                                                 
                          A.WT_PRE_QTY  IN_QTY,                                                         
                          A.WT_CURR_QTY   OUT_QTY                                                    
                   FROM TP_WT A, TD_ITEMMAS C, TD_ITEMMGRP D, TF_COMPANY E, TD_ITEMLGRP G                                                    
                    WHERE   A.LOT_NO IN     (SELECT LOT_NO FROM TP_MV WHERE MV_END_DT >= '%s' AND MV_END_DT <= '%s'  AND PRA_CD ='T010')                                                    
                        AND ( A.ITEM_CD      = C.ITEM_CD )                                                    
                        AND ( C.CMP_CD    = '002009')                                                    
                        AND ( C.ITEM_MCD     = D.ITEM_MCD )                                                    
                   AND ( C.ITEM_LCD     = G.ITEM_LCD )                                                    
                        AND ( C.CMP_CD       = E.CMP_CD )                                                    
                        AND A.WT_CURR_QTY = 0                                                    
                         AND WT_STATUS = 'PFLS'                                                    
                    GROUP BY A.PRA_CD, E.CMP_SNM, G.ITEM_LNM, D.ITEM_MNM, C.ITEM_PKG_OPT, C.PART_NO, C.SALE_CD,                                                     
                         A.LOT_NO, A.WT_PRE_QTY, A.WT_CURR_QTY                                                             
                    UNION ALL                                                    
                   SELECT 'SPOS' PRA_CD, FNP_GETCMPNM(A.CMP_CD) CMP_NAME, G.ITEM_LNM FAMILY, FNP_GETPACKAGE(A.ITEM_CD) PKG, FNP_GETCDNM(C.ITEM_PKG_OPT) PKG_OPT,                                                    
                               C.PART_NO, C.SALE_CD, FNP_GETCUSTLOTNO(A.LOT_NO) RUN_NO, A.LOT_NO, A.WT_CURR_QTY + NVL(FNP_GETLSQTY(A.LOT_NO, 'SPOS'), 0) IN_QTY, A.WT_CURR_QTY OUT_QTY                                                           
                      FROM TP_WT A, TD_ITEMMAS C, TD_ITEMLGRP G                                                      
                     WHERE A.PRA_CD = 'A050'                                                    
                        AND A.WT_STATUS = 'PFOO'                                                    
                        AND (A.ITEM_CD = C.ITEM_CD)                                                    
                        AND (C.ITEM_LCD = G.ITEM_LCD)                                                    
                        AND ( A.BSN_CD    = '*')                                                    
                        AND ( A.CMP_CD    = '002009')                                                    
                        AND (A.LOT_NO IN ((SELECT LOT_NO FROM TP_MV WHERE MV_END_DT >= '%s' AND MV_END_DT <= '%s'  AND PRA_CD ='T010')))                                                    
                        AND NOT A.PRA_CD IN (DECODE('N' , 'N', 'A050', ' '))                                                    
    )                                                    
    WHERE  PRA_CD IN DECODE('N' , 'Y', 'SPOS', PRA_CD)                                                    
    GROUP BY PRA_CD, CMP_NAME, FAMILY, PKG, PKG_OPT, PART_NO, SALE_CD, RUN_NO, LOT_NO, IN_QTY, OUT_QTY                                               

   """%(ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt,ST_dt,END_dt )
    cur.execute(sql)

    result = cur.fetchall()

    cur.close()

    con.close()

    return result