#coding:gbk
#coding:utf-8
import xlrd
excel=xlrd.open_workbook('C:/Users/Lenovo/Desktop/xlrd-1.1.0/attention.xlsx')
sheet=excel.sheet_by_name('attendlog_s')
print "  "+"学号"+"          "+"  姓名"+"       "+"             "+"    考勤汇总"
print"                                3-14    3-15     3-23     3-29     3-30"
print ""
print "201510733011"+"      "+"李茂美"+"      "+" 没有记录 没有记录 没有记录 没有记录 没有记录"
print ""
x=sheet.cell_value(1,1)
y=sheet.cell_value(1,2)
print int(x),
print "     ",
print y,
print"     ",
print" 没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(11,1)
y=sheet.cell_value(11,2)
print int(x),
print "     ",
print y,
print "     ",
print" 没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(19,1)
y=sheet.cell_value(19,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录  迟到   迟到 "
print ""
x=sheet.cell_value(23,1)
y=sheet.cell_value(23,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(27,1)
y=sheet.cell_value(27,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 没有记录 正常出勤"
print ""
x=sheet.cell_value(29,1)
y=sheet.cell_value(29,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(39,1)
y=sheet.cell_value(39,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(43,1)
y=sheet.cell_value(43,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(47,1)
y=sheet.cell_value(47,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 没有记录 正常出勤"
print ""
x=sheet.cell_value(55,1)
y=sheet.cell_value(55,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(59,1)
y=sheet.cell_value(59,2)
print int(x),
print "     ",
print y,
print "    ",
print"  正常出勤 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(70,1)
y=sheet.cell_value(70,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(74,1)
y=sheet.cell_value(74,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(78,1)
y=sheet.cell_value(78,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(82,1)
y=sheet.cell_value(82,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 正常出勤 没有记录 正常出勤 没有记录"
print ""
x=sheet.cell_value(90,1)
y=sheet.cell_value(90,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(94,1)
y=sheet.cell_value(94,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(98,1)
y=sheet.cell_value(98,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(104,1)
y=sheet.cell_value(104,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(110,1)
y=sheet.cell_value(110,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(118,1)
y=sheet.cell_value(118,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(122,1)
y=sheet.cell_value(122,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(126,1)
y=sheet.cell_value(126,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 正常记录 正常记录 没有记录 没有记录"
print ""
x=sheet.cell_value(136,1)
y=sheet.cell_value(136,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(23,1)
y=sheet.cell_value(23,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(140,1)
y=sheet.cell_value(140,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(144,1)
y=sheet.cell_value(144,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(148,1)
y=sheet.cell_value(148,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(156,1)
y=sheet.cell_value(156,2)
print int(x),
print "     ",
print y,
print "    ",
print"正常出勤 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(166,1)
y=sheet.cell_value(166,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(179,1)
y=sheet.cell_value(179,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(182,1)
y=sheet.cell_value(182,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(186,1)
y=sheet.cell_value(186,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(192,1)
y=sheet.cell_value(192,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(198,1)
y=sheet.cell_value(198,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(202,1)
y=sheet.cell_value(202,2)
print int(x),
print "     ",
print y,
print "    ",
print"  没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(212,1)
y=sheet.cell_value(212,2)
print int(x),
print "     ",
print y,
print "    ",
print"  正常出勤 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(234,1)
y=sheet.cell_value(234,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(238,1)
y=sheet.cell_value(238,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(246,1)
y=sheet.cell_value(246,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(249,1)
y=sheet.cell_value(249,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 没有记录 正常出勤"
print ""
x=sheet.cell_value(258,1)
y=sheet.cell_value(258,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(264,1)
y=sheet.cell_value(264,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(268,1)
y=sheet.cell_value(268,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤  迟到 "
print ""
x=sheet.cell_value(274,1)
y=sheet.cell_value(274,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(278,1)
y=sheet.cell_value(278,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(289,1)
y=sheet.cell_value(289,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(292,1)
y=sheet.cell_value(292,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(296,1)
y=sheet.cell_value(296,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 正常出勤 正常出勤 没有记录 没有记录"
print ""
x=sheet.cell_value(319,1)
y=sheet.cell_value(319,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 正常出勤 正常出勤"
print ""
x=sheet.cell_value(325,1)
y=sheet.cell_value(325,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 没有记录 正常出勤"
print ""
x=sheet.cell_value(331,1)
y=sheet.cell_value(331,2)
print int(x),
print "     ",
print y,
print "    ",
print"没有记录 没有记录 没有记录 没有记录 正常出勤"
print ""


    
       

