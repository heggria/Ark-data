import json
import xlwt

path = r"C:/Users/TimothyBu/Desktop/workPlace/TEST/skill_table.json"


def resolveJson(path):
    file = open(path, "rb")
    fileJson = json.load(file)

    return fileJson


def output():
    result = resolveJson(path)
    # print(result)
    title = ["技能名称", "描述", "blackboard1", "blackboard2",
             "blackboard3", "blackboard4", "blackboard5"]
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)
    flag = 0
    for index in title:
        sheet.write(0, flag, index)
        flag += 1
    flag = 1
    for index in result:
        sheet.write(flag, 0, result[index]['levels'][0]['name'])
        sheet.write(flag, 1, result[index]['levels'][0]['description'])
        sheet.write(flag, 2, result[index]['hidden'])
        sheet.write(flag, 3, result[index]['levels'][0]['rangeId'])
        sheet.write(flag, 4, result[index]['iconId'])
        sheet.write(flag, 5, result[index]['levels'][0]['skillType'])
        sheet.write(flag, 6, result[index]['levels'][0]['prefabId'])
        sheet.write(flag, 7, result[index]['levels'][0]['duration'])
        for index1 in range(len(result[index]['levels'][0]['blackboard'])):
            sheet.write(
                flag, 8+index1, result[index]['levels'][0]['blackboard'][index1]["key"])
        flag += 1
    workbook.save('demo.xls')
    print(result["skcom_charge_cost[1]"]['levels'][0]['name'])


output()
