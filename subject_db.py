from test_database import  database_proc
from test_database import  subject


class subject_db(database_proc):
    def __init__(self,subject_enum:subject):
        '''
        subject_enum ：类型为subject的枚举变量
        '''
        super().__init__()              #FIXME,init函数后面必须带括号
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
        print("######################################################")
        self.doc_parser_comm(self.test_doc_comm,self.subject_name)
        self.doc_parser(self.subject_file_path)
        self.save_database_to_excel(self.subject_name+'.xlsx','sheet1')
        print("######################################################")

    def add_new_test(self,test_type:str,test_text:str,page_num:int,importance:int,right_cnt:int=0,wrong_cnt:int=0):
        '''
        功能：向题库中加入一道新题。

        输入：
             test_type:题目类型，字符串，“简答”，“名词解释”或者“论述”
             test_text:题目内容，字符串，文字叙述
             page_num :对应书中页码，整数
             importance：重要性，0-3，默认为0
             right_cnt: 答对次数，默认为0
             wrong_cnt: 答错次数，默认为0
        '''
        tmp_dict = {self.columns_name[0]:test_type,
                    self.columns_name[1]:test_text,
                    self.columns_name[2]:page_num,
                    self.columns_name[3]:importance,
                    self.columns_name[4]:right_cnt,
                    self.columns_name[5]:wrong_cnt}

        self.database = self.database.append(tmp_dict,ignore_index=True) 

        #刷新xlsx文件
        self.save_database_to_excel(self.subject_name+'.xlsx','sheet1')

    def quiz_setter(self,type:str,num:int,importance:int=0,repeat:bool=False):
        '''
        功能：出题器
        type：出题类型
        num：题目数量
        importance：题目最低的重要性，默认为0
        repeat:是否重复，如果设置为True，则可能出现之前出过的题
        '''            
        # columns_name = ["题目类型","内容","页码","重要性","答对次数","答错次数"]
        if repeat:
            pd_tmp = self.database[(self.database["题目类型"] == type) & (self.database['重要性']>=importance)]
        else:
            pass
        pd_tmp.sample(n=num,weights=pd_tmp["重要性"])
        # print(pd_tmp)
        #len(pd_tmp.index) 获取df长度
        return pd_tmp


if __name__ == "__main__":
    # print(subject.DVP_PSY)
    subject_db_dvp = subject_db(subject.DVP_PSY)


    subject_db_dvp.get_data_form_excel(r'发展心理学.xlsx')
    print("######################################################")
    print(subject_db_dvp.database)
    print("######################################################")

    pd_des = subject_db_dvp.quiz_setter("论述",2)
    print(pd_des)
    # subject_db_dvp.add_new_test('简答','新题目',455,3)
    # print("######################################################")
    # print(subject_db_dvp.database)
    # print("######################################################")
    # subject_db_dvp.parser_subject()
    # subject_db_dvp.save_database_to_excel(r'发展心理学.xlsx',"发心")
    # print(subject_db_dvp.database)


    # subjects = [] #初始化空列表
    # for sub in subject:
    #     subjects.append(subject_db(sub))

    # for psy_sub in subjects:
    #     psy_sub.parser_subject()