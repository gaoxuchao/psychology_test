from test_database import  database_proc
from test_database import  subject


class subject_db(database_proc):
    def __init__(self,subject_enum:subject):
        '''
        subject_enum ：类型为subject的枚举变量
        '''
        super().__init__()
        #初始化题库位置
        self.test_database_path = r".\题库" "\\"

        self.sub_enum = subject_enum

        #根据传入的枚举值初始化docx文件名
        self.subject_name = subject_enum.value

        if self.sub_enum == subject.DVP_PSY:
            self.subject_path = r"发展心理学名词解释及论述.docx"
        elif self.sub_enum == subject.EDU_PSY:
            self.subject_path = r"教育心理学名词解释及论述.docx"
        elif self.sub_enum == subject.EXP_PSY:
            self.subject_path = r"实验心理学名词解释及论述.docx"     
        elif self.sub_enum == subject.COMM_PSY:
            self.subject_path = r"普通心理学名词解释及论述.docx"      
        elif self.sub_enum == subject.STA_PSY:
            self.subject_path = r"统计心理学名词解释及论述.docx"  
        elif self.sub_enum == subject.SOC_PSY:
            self.subject_path = r"社会心理学名词解释及论述.docx" 
        elif self.sub_enum == subject.MES_PSY:
            self.subject_path = r"测量名词解释及论述.docx" 
        else:
            print("Cannot find subject you enter!")

        self.subject_file_path = self.test_database_path + self.subject_path

        self.test_doc_comm     = self.test_database_path + r"心理学考研必背300题.docx"

    def parser_subject(self):
        '''
        封装父类方法，直接生成题库
        '''

        self.doc_parser_comm(self.test_doc_comm,self.subject_name)
        self.doc_parser(self.subject_file_path)



if __name__ == "__main__":
    print(subject.DVP_PSY)
    subject_db_dvp = subject_db(subject.DVP_PSY)
    subject_db_dvp.parser_subject()
    subject_db_dvp.save_database_to_excel(r'发展心理学.xlsx',"发心")
    # print(subject_db_dvp.database)