WITH CTE AS
(
--All Invoices
SELECT
A.NUM_0,A.PTHNUM_0,CASE WHEN A.POHNUM_0 = '' AND B.POHNUM_0 IS NOT NULL AND B.POHNUM_0 <> '' THEN B.POHNUM_0 ELSE A.POHNUM_0 END POHNUM_0,NULL PSHNUM_0,A.NUM_0 DOCNUM_0,12 DOCTYP_0
,ACCDAT_0,A.RCPDAT_0,C.ORDDAT_0,NULL PRQDAT_0,ACCDAT_0 DOCDAT_0
,A.TYPORI_0 INVTYPLIN_0
,A.ITMREF_0,A.NETCUR_0,A.NETPRI_0
,CASE WHEN TYPORI_0 = 3 THEN -A.QTYSTU_0 ELSE A.QTYSTU_0 END QTYSTU_0
,CASE WHEN TYPORI_0 = 3 THEN -A.AMTNOTLIN_0 ELSE A.AMTNOTLIN_0 END INVAMT_0
,CASE WHEN A.PTHNUM_0 IS NOT NULL AND A.PTHNUM_0 <> '' AND TYPORI_0 <> 3 THEN A.AMTNOTLIN_0 END RECAMT_0
,CASE WHEN A.POHNUM_0 IS NOT NULL AND A.POHNUM_0 <> '' AND TYPORI_0 <> 3 THEN A.AMTNOTLIN_0 END ORDAMT_0
,A.AMTNOTLIN_0
,0 COMAMT_0
,A.AMTNOTLIN_0 ACTAMT_0
,A.PJT_0
,A.CPY_0,A.FCYLIN_0 FCY_0,A.STU_0,A.PUU_0,A.BPSNUM_0,A.BPR_0
FROM [X3].[PROD].PINVOICED A
LEFT JOIN [X3].[PROD].PRECEIPTD B
	on A.PTHNUM_0 = B.PTHNUM_0 AND A.PTDLIN_0 = B.PTDLIN_0
LEFT JOIN [X3].[PROD].PORDERQ C
	on A.POHNUM_0 = C.POHNUM_0 AND A.POPLIN_0 = C.POPLIN_0 AND A.POQSEQ_0 = C.POQSEQ_0
WHERE A.PJT_0 <> '' AND C.POHTYP_0 = 1

UNION ALL

--All Receipts that are not all invoiced
SELECT NULL,A.PTHNUM_0,A.POHNUM_0,NULL PSHNUM_0,A.PTHNUM_0,11 DOCTYP_0
,NULL,A.RCPDAT_0,C.ORDDAT_0,NULL PRQDAT_0,A.RCPDAT_0
,NULL
,A.ITMREF_0,A.NETCUR_0,A.NETPRI_0
,CASE
	WHEN A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0) > 0
		THEN A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0)
	ELSE 0 END QTYSTU_0
,NULL INVAMT_0
,CASE WHEN A.LINAMT_0 - COALESCE(B.AMTNOTLIN_0,0) > 0 THEN A.LINAMT_0 - COALESCE(B.AMTNOTLIN_0,0) ELSE 0 END RECAMT_0
,CASE
	WHEN A.POHNUM_0 IS NOT NULL AND A.POHNUM_0 <> '' AND A.LINAMT_0 - COALESCE(B.AMTNOTLIN_0,0) > 0 THEN A.LINAMT_0 - COALESCE(B.AMTNOTLIN_0,0) ELSE 0 END ORDAMT_0
,A.LINAMT_0 - COALESCE(B.AMTNOTLIN_0,0) AMTNOTLIN_0
,0
,A.LINAMT_0 - COALESCE(B.AMTNOTLIN_0,0)
,A.PJT_0
,A.CPY_0,A.POHFCY_0,A.STU_0,A.PUU_0,A.BPSNUM_0,A.BPSINV_0
FROM [X3].[PROD].PRECEIPTD A
LEFT JOIN [X3].[PROD].PORDERQ C
	on A.POHNUM_0 = C.POHNUM_0 AND A.POPLIN_0 = C.POPLIN_0 AND A.POQSEQ_0 = C.POQSEQ_0
LEFT JOIN (SELECT PTHNUM_0,PTDLIN_0,sum(QTYSTU_0) QTYSTU_0,sum(AMTNOTLIN_0) AMTNOTLIN_0 FROM [X3].[PROD].PINVOICED WHERE PJT_0 <> '' AND PTHNUM_0 <> '' AND TYPORI_0 <> 3 GROUP BY PTHNUM_0,PTDLIN_0) B
	on A.PTHNUM_0 = B.PTHNUM_0 AND A.PTDLIN_0 = B.PTDLIN_0
WHERE A.PJT_0 <> '' AND A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0) > 0 AND A.POHTYP_0 = 1

UNION ALL

--All Orders that are not receipted
SELECT
NULL,NULL,A.POHNUM_0,NULL PSHNUM_0,A.POHNUM_0, 9 DOCTYP_0
,NULL,NULL,D.ORDDAT_0,NULL PRQDAT_0,D.ORDDAT_0
,NULL
,A.ITMREF_0,D.NETCUR_0,A.NETPRI_0
,D.QTYSTU_0 - CASE WHEN INVQTYSTU_0 >= RCPQTYSTU_0 THEN INVQTYSTU_0
						WHEN INVQTYSTU_0 <= RCPQTYSTU_0 THEN RCPQTYSTU_0
						ELSE 0 END QTYSTU_0
,NULL INVAMT_0
,NULL RECAMT_0
,D.LINAMT_0/QTYSTU_0 * CASE WHEN INVQTYSTU_0 >= RCPQTYSTU_0 THEN QTYSTU_0 - (INVQTYSTU_0)
						WHEN INVQTYSTU_0 <= RCPQTYSTU_0 THEN QTYSTU_0 - (RCPQTYSTU_0)
						ELSE 0 END ORDAMT_0
,D.LINAMT_0/QTYSTU_0 * CASE WHEN INVQTYSTU_0 >= RCPQTYSTU_0 THEN QTYSTU_0 - (INVQTYSTU_0)
						WHEN INVQTYSTU_0 <= RCPQTYSTU_0 THEN QTYSTU_0 - (RCPQTYSTU_0)
						ELSE 0 END
,0 COMAMT_0
,D.LINAMT_0/QTYSTU_0 * CASE WHEN INVQTYSTU_0 >= RCPQTYSTU_0 THEN QTYSTU_0 - (INVQTYSTU_0)
						WHEN INVQTYSTU_0 <= RCPQTYSTU_0 THEN QTYSTU_0 - (RCPQTYSTU_0)
						ELSE 0 END
,A.PJT_0
,D.CPY_0,D.POHFCY_0,D.STU_0,D.PUU_0,D.BPSNUM_0,D.BPSINV_0
FROM [X3].[PROD].PORDERP A
LEFT JOIN [X3].[PROD].PORDERQ D
	on A.POHNUM_0 = D.POHNUM_0 AND A.POPLIN_0 = D.POPLIN_0
WHERE A.PJT_0 <> '' AND D.QTYSTU_0 - CASE WHEN INVQTYSTU_0 >= RCPQTYSTU_0 THEN INVQTYSTU_0
						WHEN INVQTYSTU_0 <= RCPQTYSTU_0 THEN RCPQTYSTU_0
						ELSE 0 END > 0 AND A.POHTYP_0 = 1

UNION ALL

SELECT
NULL,NULL,NULL,A.PSHNUM_0,A.PSHNUM_0, 8 DOCTYP_0
,NULL,NULL,NULL,C.PRQDAT_0,C.PRQDAT_0
,NULL
,A.ITMREF_0,A.CUR_0,A.NETPRI_0,A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0) QTYSTU_0
,NULL
,NULL
,NULL
,A.NETPRI_0*(A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0)) AMTNOTLIN_0
,A.NETPRI_0*(A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0)) COMAMT_0
,0 ACTAMT_0
,A.PJT_0
,A.CPY_0,A.PSHFCY_0,A.STU_0,A.PUU_0,A.BPSNUM_0,NULL
FROM [X3].[PROD].PREQUISD A
LEFT JOIN [X3].[PROD].PREQUISO B
	on A.PSHNUM_0 = B.PSHNUM_0 AND A.PSDLIN_0 = B.PSDLIN_0
LEFT JOIN [X3].[PROD].PREQUIS C
	on A.PSHNUM_0 = C.PSHNUM_0
WHERE PJT_0 <> '' AND A.QTYSTU_0 - COALESCE(B.QTYSTU_0,0) <> 0 AND LINCLEFLG_0 <> 2
)


SELECT
	 CASE WHEN A.NUM_0 = '' OR A.NUM_0 IS NULL THEN ''N/A'' ELSE A.NUM_0 END NUM_0
	,CASE WHEN A.PTHNUM_0 = '' OR A.PTHNUM_0 IS NULL THEN ''N/A'' ELSE A.PTHNUM_0 END PTHNUM_0
	,CASE WHEN A.POHNUM_0 = '' OR A.POHNUM_0 IS NULL THEN ''N/A'' ELSE A.POHNUM_0 END POHNUM_0
	,CASE WHEN A.PSHNUM_0 = '' OR A.PSHNUM_0 IS NULL THEN ''N/A'' ELSE A.PSHNUM_0 END PSHNUM_0
	,A.DOCNUM_0
	,A.DOCTYP_0
	,A.ACCDAT_0
	,A.RCPDAT_0
	,A.ORDDAT_0
	,A.PRQDAT_0
	,A.DOCDAT_0
	,A.ITMREF_0
	,A.NETCUR_0
	,A.NETPRI_0
	,A.QTYSTU_0
	,A.AMTNOTLIN_0
	,A.INVAMT_0
	,A.RECAMT_0
	,A.ORDAMT_0
	,A.COMAMT_0
	,A.ACTAMT_0
	,A.PJT_0
	,B.PCCCOD_0
	,A.CPY_0
	,A.FCY_0
	,A.STU_0
	,A.PUU_0
	,A.BPSNUM_0
	,A.BPR_0
	,1 DOCFLG_0
FROM CTE A
LEFT JOIN [X3].[PROD].ITMMASTER B
	on A.ITMREF_0 = B.ITMREF_0
WHERE PCCCOD_0 <> '' AND (B.STDFLG_0 IN (4,1)
 OR (EXISTS (SELECT M.ITMREF_0
             FROM [X3].[PROD].PJMTSKITM M
				INNER JOIN [X3].[PROD].PJMTSK K
				on M.ITMREF_0 = A.ITMREF_0 AND K.KEYCONCAT_0 = A.PJT_0
				AND M.OPPNUM_0 = K.OPPNUM_0 AND M.TASCOD_0 = K.TASCOD_0)))
UNION ALL
SELECT
	 CASE WHEN A.NUM_0 = '' OR A.NUM_0 IS NULL THEN ''N/A''       ELSE C.PCCCOD2_0 || ''-'' || A.NUM_0 END NUM_0
	,CASE WHEN A.PTHNUM_0 = '' OR A.PTHNUM_0 IS NULL THEN ''N/A'' ELSE C.PCCCOD2_0 || ''-'' || A.PTHNUM_0 END PTHNUM_0
	,CASE WHEN A.POHNUM_0 = '' OR A.POHNUM_0 IS NULL THEN ''N/A'' ELSE C.PCCCOD2_0 || ''-'' || A.POHNUM_0 END POHNUM_0
	,CASE WHEN A.PSHNUM_0 = '' OR A.PSHNUM_0 IS NULL THEN ''N/A'' ELSE C.PCCCOD2_0 || ''-'' || A.PSHNUM_0 END PSHNUM_0
	,CASE WHEN A.DOCNUM_0 = '' OR A.DOCNUM_0 IS NULL THEN ''N/A'' ELSE C.PCCCOD2_0 || ''-'' || A.DOCNUM_0 END DOCNUM_0
	,A.DOCTYP_0
	,A.ACCDAT_0
	,A.RCPDAT_0
	,A.ORDDAT_0
	,A.PRQDAT_0
	,A.DOCDAT_0
	,A.ITMREF_0
	,A.NETCUR_0
	,A.NETPRI_0
	,A.QTYSTU_0
	,A.INVAMT_0*C.PCCPRCT_0/100
	,A.RECAMT_0*C.PCCPRCT_0/100
	,A.ORDAMT_0*C.PCCPRCT_0/100
	,A.AMTNOTLIN_0*C.PCCPRCT_0/100
	,A.COMAMT_0*C.PCCPRCT_0/100
	,A.ACTAMT_0*C.PCCPRCT_0/100
	,A.PJT_0
	,C.PCCCOD2_0
	,A.CPY_0
	,A.FCY_0
	,A.STU_0
	,A.PUU_0
	,A.BPSNUM_0
	,A.BPR_0
	,0 DOCFLG_0
FROM CTE A
LEFT JOIN [X3].[PROD].ITMMASTER B
	on A.ITMREF_0 = B.ITMREF_0
INNER JOIN [X3].[PROD].PJMCOSTCTR C
	on B.PCCCOD_0 = C.PCCCOD_0 AND PCCPRCT_0 <> 0
WHERE B.PCCCOD_0 <> '' AND (B.STDFLG_0 IN (4,1)
 OR (EXISTS (SELECT M.ITMREF_0
             FROM [X3].[PROD].PJMTSKITM M
				INNER JOIN [X3].[PROD].PJMTSK K
				on M.ITMREF_0 = A.ITMREF_0 AND K.KEYCONCAT_0 = A.PJT_0
				AND M.OPPNUM_0 = K.OPPNUM_0 AND M.TASCOD_0 = K.TASCOD_0)))
;