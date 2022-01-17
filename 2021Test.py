import tkinter as tk
import win32com.client
import pythoncom
import time

def vtpnt(x, y, z=0):  # Python和CAD数据类型不一样，需要转换
    """坐标点转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtobj(obj):  # Python和CAD数据类型不一样，需要转换
    """转化为对象数组"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


def vtFloat(list):  # Python和CAD数据类型不一样，需要转换
    """列表转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)


def vtInt(list):  # Python和CAD数据类型不一样，需要转换
    """列表转化为整数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)


def vtVariant(list):  # Python和CAD数据类型不一样，需要转换
    """列表转化为变体"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list)


def retryCMD(cmd: str):
    while True:
        try:
            return eval(cmd)
        except:
            time.sleep(0.001)
            print('Error at ' + cmd)
            pass


try:
    acad = win32com.client.Dispatch("AutoCAD.Application.24")
    doc = acad.ActiveDocument
    print('建立连接，文件名为: ' + doc.Name)
except:
    pass

doc = acad.ActiveDocument
mp = doc.ModelSpace
try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    pass

slt = doc.SelectionSets.Add("SS1")  # 建立一个选择集叫 SS1，默认为空集
slt.Clear()
slt.SelectOnScreen()  # 从CAD上直接点击选择我们的路线（多段线），它会自动被加入到选择集里

obj = retryCMD('slt[0]')  # 路线（多段线）这个对象
# x1, y1, x2, y2 = 461307.3822, 3408256.8258, 468316.3482, 3400883.2368
# obj = mp.AddPolyline(vtFloat([x1, y1, 0, (x1+x2)/2, (y1+y2)/2, 0, x2, y2, 0]))


number_of_interupt = 300  # 由于路线里存在圆弧，把路线在竖直方向上分成等距的多段折线，这里数字代表分成多少段，越大对于弧线的近似越好
line_width = 30  # 线路左右需要考虑的影响宽度
temp_cor = retryCMD('list(obj.Coordinates)')
# 路线起点和终点坐标
x_up, y_up, x_low, y_low = temp_cor[0], temp_cor[1], temp_cor[-2], temp_cor[-1]
interval_y = (y_up - y_low) / number_of_interupt
interval_x = (x_up - x_low) / number_of_interupt

new_center_coor = []  # 记录转换为多段折线（不含弧线）的路线坐标
direction = '竖'
for i in range(number_of_interupt):
    if direction == '竖':
        interupt_line = mp.AddLine(
            vtpnt(x_up, y_up - i * interval_y), vtpnt(x_up + 10, y_up - i * interval_y))
    else:
        interupt_line = mp.AddLine(
            vtpnt(x_up - i * interval_x, y_up), vtpnt(x_up - i * interval_x, y_up + 10))
    temp = retryCMD('list(obj.IntersectWith(interupt_line, 3))')
    new_center_coor.append(temp[0])
    new_center_coor.append(temp[1])
    new_center_coor.append(temp[2])
    retryCMD('interupt_line.Delete()')

new_center_coor.append(x_low)
new_center_coor.append(y_low)
new_center_coor.append(0)

obj_new = mp.AddPolyline(vtFloat(new_center_coor))  # 折线化处理后的路线对象

temp = obj_new.Offset(str(line_width))  # 路线左右偏移
offset1 = temp[0]  # 偏移线1

temp = obj_new.Offset(str(-line_width))  # 路线左右偏移
offset2 = temp[0]  # 偏移线2

# nvetex = number_of_interupt+1 #路线节点个数
nvetex1 = int(len(list(offset1.Coordinates)) / 3)  # 路线节点个数
nvetex2 = int(len(list(offset2.Coordinates)) / 3)  # 路线节点个数
lineoffset1 = list(offset1.Coordinates)  # 路线左偏移后的所有节点坐标
lineoffset2 = list(offset2.Coordinates)  # 路线右偏移后的所有节点坐标

cors = []  # 将左偏移线和右偏移线连成一个闭合多段线，cors是这个多段线的所有节点坐标
for i in range(nvetex1):
    cors.insert(0, 0)
    cors.insert(0, lineoffset1[3 * (nvetex1 - 1 - i) + 1])
    cors.insert(0, lineoffset1[3 * (nvetex1 - 1 - i)])

for i in range(nvetex2):
    cors.append(lineoffset2[3 * (nvetex2 - 1 - i)])
    cors.append(lineoffset2[3 * (nvetex2 - 1 - i) + 1])
    cors.append(0)

cors.append(cors[0])
cors.append(cors[1])
cors.append(cors[2])

acad.Update()
slt.Clear()  # 清空选择集
# 用左右偏移线连成的闭合多段线去窗交建筑物，并加入选择集
retryCMD('slt.SelectByPolygon(2, vtFloat(cors))')
# 用左右偏移线连成的闭合多段线去框选建筑物，并加入选择集
retryCMD('slt.SelectByPolygon(6, vtFloat(cors))')

sum_area = 0  # 统计总面积
n_blocks = retryCMD('slt.Count')
blocks = []
for i in range(n_blocks):
    block = retryCMD('slt[i]')
    if block not in [obj, obj_new, offset1, offset2]:
        block.Color = 2  # 选择集里的对象高亮
        blocks.append(block)
        sum_area += float(block.Area)  # 面积累加

rc = mp.AddPolyline(vtFloat(cors))
output = "房屋总面积为 %.1f 平米\n" % (sum_area)
print(output)
acad.Update()

time.sleep(5)

for i in blocks:
    i.Color = 255
# obj.Delete()
obj_new.Delete()
offset1.Delete()
offset2.Delete()
rc.Delete()
acad.Update()
