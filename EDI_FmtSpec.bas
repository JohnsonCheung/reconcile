Attribute VB_Name = "EDI_FmtSpec"
Option Compare Database
Private Sub DE1()
'AliasH | FldNm              | Alias
'AliasL | DESADV Number      | DENo
'AliasL | DESADV Line Number | DELNo
'AliasL | ORDER Number       | OrdNo
'AliasL | ORDER Line Number  | OrdLNo
'AliasL | Item Number        | ItmNo
'AliasL | Item Description   | ItmDes
'AliasL | Quantity           | Qty
'AliasL | Unit of Measure    | UOM
'AliasL | Lot Number         | Lot
'AliasL | Retest Date        | TstDte
'AliasL | Expiry Date        | ExpDte
'AliasL | Manufactor Lot     | MLot
'AliasL | Customs Code       | CusmCd
'FreezeH | Address
'FreezeL | I4
'DteH | AliasLvs
'DteL | TstDte ExpDte
'NumH | AliasLvs
'NumL | DENo DELNo OrdNo OrdLNo Qty
'ColrH | ColrNm | AliasLvs
'ColrL | *Green | Qty
'HdrHgtH | Factor
'HdrHgtL | 3
'YYYYMMDDH | AliasLvs
'YYYYMMDDL | ExpDte TstDte
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | Qty OrdNo OrdLNo DENo DELNo ExpDte TstDte UOM CusmCd
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,##0     | Qty
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ItmDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | Qty
'TotL  | *Cnt    | DESH
'WdtH | Wdt | ColNmLvs
'WdtL | 10  | ExpDte TstDte UOM OrdNo DENo ItmNo DELNo
'WdtL | 8   | OrdLNo UOM
'WdtL | 6   | DESH
'WdtL | 9   | Qty
End Sub
Private Sub HANMOV()
'AliasH | FldNam                | Alias
'AliasL | HANMOV Number         | HMNo
'AliasL | HANMOV Line Number    | HMLNo
'AliasL | ORDER Number          | OrdNo
'AliasL | ORDER Line Number     | OrdLNo
'AliasL | Item Number           | ItmNo
'AliasL | Item Description      | ItmDes
'AliasL | Quantity              | Qty
'AliasL | Lot Number            | Lot
'AliasL | Unit of Measure       | UOM
'AliasL | Customs Code          | CusmCd
'AliasL | Ingrediend Labels     | IngLbl
'AliasL | Pre -Printed          | PrePrt
'AliasL | Date Label            | DteLbl
'AliasL | Importer name label   | ImpLbl
'AliasL | Extra Label           | ExtLbl
'AliasL | Labelling instruction | LblInst
'FreezeH | Address
'FreezeL | I4
'DteH | AliasLvs
'NumH | AliasLvs
'NumL | HMNo HMLNo OrdNo OrdLNo Qty
'ColrH | ColrNm | AliasLvs
'ColrL | *Green | Qty
'HdrHgtH | Factor
'HdrHgtL | 3
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | Qty OrdNo OrdLNo HMNo HMLNo UOM CusmCd
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,##0     | Qty
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ItmDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | Qty
'TotL  | *Cnt    | HANH
'WdtH | Wdt | ColNmLvs
'WdtL | 10  | UOM OrdNo HMNo ItmNo HMLNo
'WdtL | 8   | OrdLNo UOM
'WdtL | 6   | HANH
'WdtL | 9   | Qty
End Sub
Private Sub DE2()
'AliasH | FldNm              | Alias
'AliasL | DESADV Number      | DENo
'AliasL | DESADV Line Number | DELNo
'AliasL | ORDER Number       | OrdNo
'AliasL | ORDER Line Number  | OrdLNo
'AliasL | Item Number        | ItmNo
'AliasL | Item Description   | ItmDes
'AliasL | Quantity           | Qty
'AliasL | Unit of Measure    | UOM
'AliasL | Lot Number         | Lot
'AliasL | Retest Date        | TstDte
'AliasL | Expiry Date        | ExpDte
'AliasL | Manufactor Lot     | MLot
'AliasL | Customs Code       | CusmCd
'FreezeH | Address
'FreezeL | I4
'DteH | AliasLvs
'DteL | TstDte ExpDte
'NumH | AliasLvs
'NumL | DENo DELNo OrdNo OrdLNo Qty
'ColrH | ColrNm | AliasLvs
'ColrL | *Green | Qty
'HdrHgtH | Factor
'HdrHgtL | 3
'YYYYMMDDH | AliasLvs
'YYYYMMDDL | ExpDte TstDte
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | Qty OrdNo OrdLNo DENo DELNo ExpDte TstDte UOM CusmCd
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,##0     | Qty
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ItmDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | Qty
'TotL  | *Cnt    | DESH
'WdtH | Wdt | ColNmLvs
'WdtL | 10  | ExpDte TstDte UOM OrdNo DENo ItmNo DELNo
'WdtL | 8   | OrdLNo UOM
'WdtL | 6   | DESH
'WdtL | 9   | Qty
End Sub
Private Sub IRP()
'AliasH | FldNam            | Alias
'AliasL | Warehouse         | Whs
'AliasL | Seq Number        | SeqNo
'AliasL | Item  Number      | ItmNo
'AliasL | Item  Description | ItmDes
'AliasL | Quantity          | Qty
'AliasL | Unit of Measure   | UOM
'AliasL | Lot Number        | Lot
'AliasL | Retest Date       | TstDte
'AliasL | Expiry Date       | ExpDte
'AliasL | Stock Status      | Sts
'AliasL | Base Date         | BasDte
'AliasL | Reason            | Reas
'FreezeH | Address
'FreezeL | F4
'DteH | AliasLvs
'NumH | AliasLvs
'NumL | SeqNo Qty
'ColrH | ColrNm | AliasLvs
'ColrL | *Green | Qty
'HdrHgtH | Factor
'HdrHgtL | 3
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | Whs SeqNo UOM TstDte ExpDte Sts BasDte Reas
'HAlignL | *Right     | Qty
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,##0     | Qty
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ItmDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | Qty
'TotL  | *Cnt    | INVH
'WdtH | Wdt | ColNmLvs
'WdtL | 10  | UOM ItmNo
'WdtL | 5   | Sts
'WdtL | 8   | SeqNo UOM
'WdtL | 7   | INVH Whs
'WdtL | 10  | TstDte ExpDte BasDte
'YYYYMMDDH | AliasLvs
'YYYYMMDDL | ExpDte TstDte BasDte

End Sub
Private Sub SPO()
'AliasH | FldNam                      | Alias
'AliasL | Seq. Nbr                    | SeqNo
'AliasL | Item Nbr                    | ItmNo
'AliasL | Item Description            | ItmDes
'AliasL | Req. Net Quantity           | QReq
'AliasL | Estimated Scrap             | EstScp
'AliasL | Unit of Measure             | UOM
'AliasL | Weight per Candy            | Wgt
'AliasL | Replace Item                | IsReplItm
'AliasL | Lot Ctrl                    | IsLotCtl
'AliasL | Stock Item                  | IsStkItm
'AliasL | Exclude from Assembly       | IsExclAss
'AliasL | Scrap Quantity              | QScp
'AliasL | LOT1 Nbr                    | L1
'AliasL | Quantity Consumed from LOT1 | Q1
'AliasL | LOT2 Nbr                    | L2
'AliasL | Quantity Consumed from LOT2 | Q2
'AliasL | LOT3 Nbr                    | L3
'AliasL | Quantity Consumed from LOT3 | Q3
'AliasL | LOT4 Nbr                    | L4
'AliasL | Quantity Consumed from LOT4 | Q4
'FreezeH | Address
'FreezeL | F4
'DteH | AliasLvs
'NumH | AliasLvs
'NumL | SeqNo EstScp Wgt QScp QReq Q1 Q2 Q3 Q4
'ColrH | ColrNm | AliasLvs
'HdrHgtH | Factor
'HdrHgtL | 5
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | EstScp QScp UOM Wgt IsReplItm IsStkItm IsExclAss IsLotCtl SeqNo
'HAlignL | *Right     | Q1 Q2 Q3 Q4 QReq
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,###     | Q1 Q2 Q3 Q4 QScp EstScp Wgt
'NumFmtL | #,##0     | QReq
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ItmDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | Q1 Q2 Q3 Q4 EstScp QReq QScp
'TotL  | *Cnt    | BOMH
'WdtH | Wdt | ColNmLvs
'WdtL | 10  | UOM ItmNo
'WdtL | 5   | IsLotCtl
'WdtL | 8   | SeqNo IsReplItm UOM Wgt
'WdtL | 9   | IsExclAss
'WdtL | 7   | BOMH IsStkItm
'WdtL | 10  | Q1 Q2 Q3 Q4 QScp QReq EstScp
End Sub

Private Sub IMN()
'AliasH | FldNam                     | Alias
'AliasL | Line Number                | LNo
'AliasL | ShipTo Code                | ShpTo
'AliasL | Warehouse Code             | Whs
'AliasL | Purchase Order Number      | PONo
'AliasL | Purchase Order Line Number | POLNo
'AliasL | Planned Dock Date          | DckDte
'AliasL | Supplier Name              | SupNm
'AliasL | Supplier Identifier        | SupId
'AliasL | Product Identifier         | PrdId
'AliasL | Product Description        | PrdDes
'AliasL | Expected Quantity          | QExp
'AliasL | Lot Control Flag           | IsLotCtl
'AliasL | Unit of Measure            | UOM
'AliasL | Received Quantity          | QRec
'AliasL | Supplier Lot Number        | SLot
'AliasL | Assigned Lot Number        | ALot
'AliasL | Production Date            | PrdDte
'AliasL | Expiry Date                | ExpDte
'AliasL | Comment                    | Cmt
'FreezeH | Address
'FreezeL | M4
'DteH | AliasLvs
'NumH | AliasLvs
'NumL | LNo PONo POLNo QExp QRec
'ColrH | ColrNm | AliasLvs
'ColrL | *Green | QRec
'HdrHgtH | Factor
'HdrHgtL | 5
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | PONo POLNo LNo UOM Whs SupId PrdDte ExpDte DckDte IsLotCtl
'HAlignL | *Right     | QExp QRec
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,##0     | QRec QExp
'YYYYMMDDH | AliasLvs
'YYYYMMDDL | DckDte PrdDte ExpDte
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ShpTo Whs SupNm SupId PrdDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | QExp QRec
'TotL  | *Cnt    | IMNH
'WdtH | Wdt | ColNmLvs
'WdtL | 7   | IMNH ShpTo
'WdtL | 8   | LNo
'WdtL | 9   | IsLotCtl POLNo
'WdtL | 10  | UOM PrdId PONo DckDte ExpDte QRec QExp SupId
'WdtL | 11  | Whs PrdDte
'WdtL | 12  | ALot SLot

End Sub

Private Sub IVM()

End Sub

Private Sub LPD()
'AliasH | FldNam                      | Alias
'AliasL | Seq. Nbr                    | SeqNo
'AliasL | Item Nbr                    | ItmNo
'AliasL | Item Description            | ItmDes
'AliasL | Req. Net Quantity           | QReq
'AliasL | Estimated Scrap             | QScpEst
'AliasL | Unit of Measure             | UOM
'AliasL | Weight per Candy            | Wgt
'AliasL | Replace Item?               | IsReplItm
'AliasL | Lot Ctrl?                   | IsLotCtl
'AliasL | Stock Item?                 | IsStkItm
'AliasL | Exclude from Assembly       | IsExclAss
'AliasL | Scrap Quantity              | QScp
'AliasL | LOT1 Nbr                    | L1
'AliasL | Quantity Consumed from LOT1 | Q1
'AliasL | LOT2 Nbr                    | L2
'AliasL | Quantity Consumed from LOT2 | Q2
'AliasL | LOT3 Nbr                    | L3
'AliasL | Quantity Consumed from LOT3 | Q3
'AliasL | LOT4 Nbr                    | L4
'AliasL | Quantity Consumed from LOT4 | Q4
'FreezeH | Address
'FreezeL | F4
'DteH | AliasLvs
'NumH | AliasLvs
'NumL | SeqNo QScpEst Wgt QScp QReq Q1 Q2 Q3 Q4
'ColrH | ColrNm | AliasLvs
'HdrHgtH | Factor
'HdrHgtL | 5
'HAlignH | HAlignType | AliasLvs
'HAlignL | *Center    | QScpEst QScp UOM Wgt IsReplItm IsStkItm IsExclAss IsLotCtl SeqNo
'HAlignL | *Right     | Q1 Q2 Q3 Q4 QReq
'NumFmtH | NumFmtStr | AliasLvs
'NumFmtL | #,###     | Q1 Q2 Q3 Q4 QScp QScpEst Wgt
'NumFmtL | #,##0     | QReq
'OutlineH | Level | AliasLvs
'OutlineL | 2     | ItmDes
'TotH  | TotType | AliasLvs
'TotL  | *Tot    | Q1 Q2 Q3 Q4 QScpEst QReq QScp
'TotL  | *Cnt    | BOMH
'WdtH | Wdt | ColNmLvs
'WdtL | 10  | UOM ItmNo
'WdtL | 5   | IsLotCtl
'WdtL | 8   | SeqNo IsReplItm UOM Wgt
'WdtL | 9   | IsExclAss
'WdtL | 7   | BOMH IsStkItm
'WdtL | 10  | Q1 Q2 Q3 Q4 QScp QReq QScpEst
End Sub

Private Sub PMU()

End Sub
