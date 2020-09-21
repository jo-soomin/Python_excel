import cx_Oracle
import sys

def item_query_P():

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
                WHERE  ( A.MV_END_DT >= '20200828080000' )
                    AND ( A.MV_END_DT <= '20200904080000' )
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
                            AND (A.WT_ACT_DT >= '20200828080000')
                            AND (A.WT_ACT_DT <  '20200904080000')
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
                WHERE  ( A.WT_ACT_DT >= '20200828080000' )
                    AND ( A.WT_ACT_DT <= '20200904080000' )
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
                WHERE  ( A.MV_END_DT >= '20200828080000' )
                    AND ( A.MV_END_DT <= '20200904080000'   )
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
   """
    cur.execute(sql)

    result = cur.fetchall()

    cur.close()

    con.close()

    return result
