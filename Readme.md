# 概述  
团队使用Metersphere管理用例有以下困难：  
1. 现有用例需要导入Metersphere。现有用例包括Excel格式和Word格式。  
   解决方案：  
   Excel格式用例转换成Metersphere导入格式(Excel)较为简单，有时仅涉及列的映射，调整列名即可。较为复杂的情况可写脚本转换，由于转换脚本不通用，不放在当前Repository里；    
   Word格式用例转换成Metersphere导入格式，需识别Word用例中各字段，参见下方“Word用例转MS导入格式”
2. 用例导出为固定格式，但是客户项目的用例根据客户需求有不同格式，需要做转换：  
   解决方案：  
   交付用例为Excel列表格式的（即和Metersphere导出用例格式一样，在单个Sheet上以列表形式列出用例的），和原有Excel格式用例转换成Metersphere导入用例格式类似，调整列名、列序，或编写转换脚本（脚本不通用，此处忽略）  
   交付用例为Excel独立sheet格式的（即每个sheet一个单独的用例表格），参见下方"生成独立sheet用例"  
   交付用例为Word格式的，参加下方“生成Word格式用例”
   

# Word用例转MS导入格式  
1. 将现有Word用例去除其他部分(如封面页、文档描述等)，仅保留用例章节，单独保存一份文档（docx格式）——参考existing_word_tc.docx（文档内容做了些脱敏处理，仅作示意） 
2. 修改tc_fields_loc中各字段对应值在Word用例表格中的位置(行,列),以0起始  
3. 修改root_dir，source_path，dest_wb_path  
4. 设置tc_steps_prefix和expected_result_prefix（去除前缀后为需要提取的内容）  
5. 运行脚本tc_word_to_excel.py
6. 将Excel文件导入Metersphere平台  
对于原有Word用例中保护执行结果截图的情况，如需将截图上传至Metersphere，当前采用的方法是在Metersphere中的用例模板中添加一个富文本字段，比如“附件字段”，将embeded_images_path对应路径下生成的图片批量使用root权限上传到metersphere服务器/app/metersphere/data/image/markdown路径下，然后将第5步生成的Excel用例导入Metersphere平台（图片插入到用例的“附加字段”字段）


# 生成独立sheet用例  
1. Metersphere中按顺序（如用例ID）排列全部用例，全部导出，勾选全部基础字段、自定义字段，以及其他字段中的评论（需要带截图时须勾选此项）、创建人、创建时间（排序用）、更新时间等
2. 检查导出的Excel用例的顺序，如果和预期输出的用例顺序不一致，手动调整（也可在生成Excel后调整Excel文档中用例顺序）
3. 修改rootdir （脚本中使用当前文件路径，可修改为客户项目文件夹)，rootdir下创建好tmpl_wb_path指向的用例模板文件，参加Excel分sheet用例模板.xlsx
4. 设置src_wb_path
5. 设置tmpl_sheet_name和tmpl_sheet_name
6. 修改actual_result_loc（实际结果内容在表格中的位置索引，如A7）
7. 设置need_image(若输出的用例不需要带截图，则设置为False)。需要输出截图时注意设置img_width和img_height
8. 修改image_dir(用例截图所在路径，服务器上为固定地址，仅调试脚本时需要修改)
9. 运行脚本populate_excel_sheets.py  
注意如果导出用例需要附带截图，则本脚本需在Metersphere所在服务器执行（否则找不到图片）；  
如果不需要附带截图，则脚本可在本地执行


#  生成Word格式用例  
生成Word格式用例相对复杂，本文仅给出总体步骤，细节见对应脚本注释  
1. 从Metersphere中导出用例，参见同“生成独立sheet用例”步骤1
2. 检查导出的Excel用例的顺序，如果和预期输出的用例顺序不一致，手动调整
3. 将客户的WORD格式用例模板裁剪到仅保留章节层级和一张用例表格，例如existing_word_tc.docx文件的格式裁剪为Testcase_template.docx模板文件。最终的用例包含需要的所有层级，例如最低层级为1.1.1.1，模板中需要包含1.1,1.1.1,1.1.1.1三个层级，和一个用例表格，如下：  
<img src=".\mdimg\template.png"  width = 400 height = 500>  
4. 插入域: 脚本使用Word的邮件合并功能来根据模板生成Word文档，因此需在对应单元格内插入域。例如上面的表格，在测试人后的空白单元格插入域（MergeField/合并域）：  
<img src=".\mdimg\insert_field1.png">  
<img src=".\mdimg\insert_field2.png">
5. 注意域名需要与脚本中对应字段名一致，以下是域名的描述（不同项目模板中，表格内字段名可能不同）：  
"test_case_id"：用例编号,   
"test_case_name": 用例名称,  
"precondition": 前置条件,  
"test_purpose": 测试目的,  
"test_steps": 测试步骤,  
"expected_result": 预期结果,  
"label": 标签,  
"test_designer": 设计人,  
"test_executor": 执行人,  
"tc_level": 用例等级,  
"tc_status": 用例状态,  
"tc_type": 用例类型,  
"test_date": 测试日期  
得到如下的结果（“实际结果”字段不使用合并域替换，因为如果带有结果，一次只能合并一张图片）：  
<img src=".\mdimg\fields_inserted.png">  
6. 生成WORD用例  
使用populate_word_mailmerge.py脚本生成Word格式的用例**<span style="background-color:yellow;">(注意需安装docx-mailmerge, 不是mailmerge，且安装有mailmerge时不能正确从docx-mailmerge导入所需的模块) </span>** 
修改脚本描述中列出的字段值，执行脚本  
如果用例中需包含截图，脚本需在Metersphere服务器端执行脚本(/home/appadmin/QATest/metersphere)  







