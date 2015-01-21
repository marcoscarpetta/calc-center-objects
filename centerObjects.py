import uno

ctx = uno.getComponentContext()
smgr = ctx.ServiceManager

def centerObjects():
  
  desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
  doc = desktop.getCurrentComponent()
  sheet = doc.CurrentController.getActiveSheet()
  c1 = sheet.createCursor()
  c1.gotoStartOfUsedArea(False)
  c1.gotoEndOfUsedArea(True)
  merged = []
  for i in range(c1.RangeAddress.EndColumn+1):
    for j in range(c1.RangeAddress.EndRow+1):
      if sheet.getCellByPosition(i,j).IsMerged:
        c = sheet.createCursorByRange(sheet.getCellRangeByPosition(i,j,i,j))
        c.collapseToMergedArea()
        v = c.RangeAddress
        merged.append([v.StartRow, v.EndRow, v.StartColumn, v.EndColumn])
  
  #raise Exception(str(merged))
  
  i=0
  while i < sheet.getDrawPage().Count:
    obj = sheet.getDrawPage().getByIndex(i)
    if hasattr(obj.Anchor, "CellAddress"):
      r = obj.Anchor.CellAddress.Row
      c = obj.Anchor.CellAddress.Column
      rowHeight = sheet.Rows.getByIndex(r).Height
      colWidth = sheet.Columns.getByIndex(c).Width
      x = sheet.Columns.getByIndex(c).Position.X
      y = sheet.Rows.getByIndex(r).Position.Y
      
      for m in merged:
        if m[0]<=r<=m[1] and m[2]<=c<=m[3]:
          rowHeight = 0
          colWidth = 0
          x = sheet.Columns.getByIndex(m[2]).Position.X
          y = sheet.Rows.getByIndex(m[0]).Position.Y
          n = m[0]
          while n <= m[1]:
            rowHeight += sheet.Rows.getByIndex(n).Height
            n += 1
          n = m[2]
          while n <= m[3]:
            colWidth += sheet.Columns.getByIndex(n).Width
            n += 1
      p = obj.Position
      p.X = x+(colWidth-obj.Size.value.Width)/2
      p.Y = y+(rowHeight-obj.Size.value.Height)/2
      obj.setPosition(p)
    i += 1