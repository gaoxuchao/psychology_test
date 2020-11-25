import subject_db
from test_database import subject
import pandas as pd
import numpy as np
from docx import Document

class main_gui() :

    def __init__(self):

        self.subject_list = [] # 创建空列表

        subject_id = 0
        for sub in subject: # 在枚举类型中遍历
            # 初始化各个科目的对象
            self.subject_list.append(subject_db.subject_db(sub,subject_id))
            subject_id +=1

        for psy_sub in self.subject_list:
            # 获取各个科目的数据
            psy_sub.get_data_form_excel(psy_sub.subject_name + '.xlsx')
            # print(psy_sub.database)
            # psy_sub.database["科目"] = psy_sub.subject_name

            # print(psy_sub.database)
            # tmp_idx = psy_sub.database.index.astype('str')
            
            # psy_sub.database["科目"] = psy_sub.database["科目"] + tmp_idx
            # psy_sub.database.set_index("科目",drop=False,inplace=True)
    

    def make_test(self,mark_db:bool=False):
        for i in range(len(self.subject_list)):
            if (i == 0): #init test 
                self.des_set      = self.subject_list[i].auto_quiz_setter_des()
                self.sim_des_set  = self.subject_list[i].auto_quiz_setter_sim_des()
                self.term_expl_set= self.subject_list[i].auto_quiz_setter_term_expl()
                self.des_set['科目']     = self.subject_list[i].subject_name
                self.des_set['科目编号'] = self.subject_list[i].subject_id 
                self.sim_des_set['科目']     = self.subject_list[i].subject_name
                self.sim_des_set['科目编号'] = self.subject_list[i].subject_id 
                self.term_expl_set['科目']     = self.subject_list[i].subject_name
                self.term_expl_set['科目编号'] = self.subject_list[i].subject_id 
            else:
                tmp_dp0 = self.subject_list[i].auto_quiz_setter_des()
                tmp_dp1 = self.subject_list[i].auto_quiz_setter_sim_des()
                tmp_dp2 = self.subject_list[i].auto_quiz_setter_term_expl() 
                tmp_dp0['科目']     = self.subject_list[i].subject_name
                tmp_dp0['科目编号'] = self.subject_list[i].subject_id 
                tmp_dp1['科目']     = self.subject_list[i].subject_name
                tmp_dp1['科目编号'] = self.subject_list[i].subject_id 
                tmp_dp2['科目']     = self.subject_list[i].subject_name
                tmp_dp2['科目编号'] = self.subject_list[i].subject_id 
                self.des_set      = self.des_set.append      ( tmp_dp0)
                self.sim_des_set  = self.sim_des_set.append  ( tmp_dp1)
                self.term_expl_set= self.term_expl_set.append( tmp_dp2) 

    def save_test(self,test_no:int):
        frames = [self.des_set,self.sim_des_set,self.term_expl_set]
        
        self.test_db = pd.concat(frames)

        self.test_db["题目索引"] = self.test_db.index

        self.test_db["正确"] = 0

        del self.test_db['内容']

        self.test_db["index"] = range(len(self.test_db))
        self.test_db = self.test_db.set_index(['index'])
        
        file_name = "试卷" + str(test_no) + '.xlsx'
        
        with pd.ExcelWriter(file_name) as writer: # pylint: disable=abstract-class-instantiated
            self.test_db.to_excel(writer,'sheet',index=True,index_label="index") #将pd写入example_file，索引列名为题号

    def get_test_from_excel(self,test_no:int):
        '''
           从EXCEL中获取之前出过的题目
        '''
        # 不要和test_db冲突
        # self.test_db_old=
        file_name = "试卷" + str(test_no) + '.xlsx'
        with pd.ExcelFile(file_name) as xls:
            sheet_names = xls.sheet_names  #获取excel文件的sheet列表
            self.test_db_old = pd.read_excel(xls,sheet_names[0],index_col=0)  #指定第一列为索引列      


    def test_feedback(self,test_no:int):
        '''
        反馈答题结果：
           从EXCEL中获取答题结果，然后根据科目，科目编号和题目索引定位到具体题目
           根据正确与否，对答对次数和答错次数进行修改
        '''
        self.get_test_from_excel(test_no)
        pass



    def init_xlxs_from_docs(self):
        for psy_sub in self.subject_list:
            psy_sub.parser_subject()

    def create_test_doc(self,test_no:int):
        test_doc = Document()
        test_name = "试卷" + str(test_no)

        test_doc.add_heading(test_name)

        test_doc.add_heading("名词解释 (3分)",level=1)
        for i in range(len(self.term_expl_set.index)):
            no_tmp = i +1
            dp_tmp = self.term_expl_set.iloc[i]
            content_tmp = dp_tmp["内容"]
            page_num_tmp = dp_tmp["页码"]
            para_tmp = str(no_tmp) +". " + content_tmp + "【" + str(page_num_tmp) + "】"
            p_term = test_doc.add_paragraph(para_tmp)

        test_doc.add_heading("简答 (15分)",level=1)
        for i in range(len(self.sim_des_set.index)):
            no_tmp = i +1
            dp_tmp = self.sim_des_set.iloc[i]
            content_tmp = dp_tmp["内容"]
            page_num_tmp = dp_tmp["页码"]
            para_tmp = str(no_tmp) +". " + content_tmp + "【" + str(page_num_tmp) + "】"
            p_term = test_doc.add_paragraph(para_tmp)

        test_doc.add_heading("论述 (25分)",level=1)
        for i in range(len(self.des_set.index)):
            no_tmp = i +1
            dp_tmp = self.des_set.iloc[i]
            content_tmp = dp_tmp["内容"]
            page_num_tmp = dp_tmp["页码"]
            para_tmp = str(no_tmp) +". " + content_tmp + "【" + str(page_num_tmp) + "】"
            p_term = test_doc.add_paragraph(para_tmp)

        test_doc.add_page_break()


        test_doc.save(test_name + ".docx")

                     
if __name__ == "__main__":
    test = main_gui()
    # test.init_xlxs_from_docs()
    test.make_test()
    test.save_test(0)
    # test.create_test_doc(4)
    print("######################################################")
    print(test.test_db)
    print("######################################################")
