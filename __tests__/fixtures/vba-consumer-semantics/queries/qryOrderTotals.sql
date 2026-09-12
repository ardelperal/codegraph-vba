SELECT l.OrderId, p.ProductName
FROM tblOrderLines AS l
INNER JOIN tblProducts AS p ON l.ProductId = p.ProductId
