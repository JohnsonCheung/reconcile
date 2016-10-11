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
'AliasH | FldNam | Alias
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

End Sub
Private Sub IRP()

End Sub
Private Sub SPO()

End Sub
Private Sub IMN()

End Sub

