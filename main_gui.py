import subject_db
from test_database import subject
from docx import Document

class main_gui() :

    def __init__(self):

        self.subject_list = [] # 创建空列表

        for sub in subject: # 在枚举类型中遍历
            # 初始化各个科目的对象
            self.subject_list.append(subject_db.subject_db(sub))

        for psy_sub in self.subject_list:
            # 获取各个科目的数据
            psy_sub.get_data_form_excel(psy_sub.subject_name + '.xlsx')
            # print(psy_sub.database)
            psy_sub.database["科目"] = psy_sub.subject_name

            # print(psy_sub.database)
            tmp_idx = psy_sub.database.index.astype('str')
            
            psy_sub.database["科目"] = psy_sub.database["科目"] + tmp_idx
            psy_sub.database.set_index("科目",drop=False,inplace=True)
    

    def make_test(self):
        for i in range(len(self.subject_list)):
            if (i == 0): #init test 
                self.des_set      = self.subject_list[i].auto_quiz_setter_des()
                self.sim_des_set  = self.subject_list[i].auto_quiz_setter_sim_des()
                self.term_expl_set= self.subject_list[i].auto_quiz_setter_term_expl()
            else:
                tmp_dp0 = self.subject_list[i].auto_quiz_setter_des()
                tmp_dp1 = self.subject_list[i].auto_quiz_setter_sim_des()
                tmp_dp2 = self.subject_list[i].auto_quiz_setter_term_expl() 
                self.des_set      = self.des_set.append      ( tmp_dp0)
                self.sim_des_set  = self.sim_des_set.append  ( tmp_dp1)
                self.term_expl_set= self.term_expl_set.append( tmp_dp2) 


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
    test.create_test_doc(1)
    print("######################################################")
    print("######################################################")
    # print(test.sim_des_set)
