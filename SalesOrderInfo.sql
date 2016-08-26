SELECT Rel.OurReqQty AS [QTY Owed], Rel.OurStockQty AS [Stock],
       SO.DocUnitPrice AS [$/Per], ROUND(Rel.OurReqQty * SO.DocUnitPrice, 2) AS [Ext. Price]
FROM Erp.OrderRel Rel
INNER JOIN Erp.OrderDtl SO ON
Rel.Company = SO.Company AND
Rel.OrderNum = SO.OrderNum AND
Rel.OrderLine = SO.OrderLine
WHERE Rel.OrderNum = '75039' AND
      Rel.OrderLine = '1'    AND
      Rel.OrderRelNum = '1'