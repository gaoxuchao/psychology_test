from test_database import  database_proc
from test_database import  subject


class subject_db(database_proc):
    def __init__(self,subject_enum:subject,sub_id:int):
        '''
        subject_enum ：类型为subject的枚举变量
        sub_id       : 顶层传入的科目编号，用于在多个科目中定位
        '''
        super().__init__()              #FIXME,init函数后面必须带括号
        #初始化题库位置
        self.test_database_path = r".\题库" "\\"

        self.sub_enum = subject_enum

        self.subject_id = sub_id

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

        self.sub_xlsx_file_name = self.test_database_path + self.subject_name

        self.sim_des_impt  = 1 #简答题出题难度默认最低为1
        self.des_impt      = 1 #论述题出题难度默认最低为1
        self.term_expl_impt= 1 #名词解释出题难度默认最低为1

        self.test_num_init()

    def parser_subject(self):
        '''
        封装父类方法，直接生成题库
        '''
        print("######################################################")
        self.doc_parser_comm(self.test_doc_comm,self.subject_name)
        self.doc_parser(self.subject_file_path)
        self.save_database_to_excel(self.sub_xlsx_file_name+'.xlsx','sheet1')
        print("######################################################")

    def add_new_test(self,test_type:str,test_text:str,page_num:int,importance:int=1,right_cnt:int=0,wrong_cnt:int=0):
        '''
        功能：向题库中加入一道新题。

        输入：
             test_type:题目类型，字符串，“简答”，“名词解释”或者“论述”
             test_text:题目内容，字符串，文字叙述
             page_num :对应书中页码，整数
             importance：重要性，0-3，默认为1
             right_cnt: 答对次数，默认为0
             wrong_cnt: 答错次数，默认为0
        '''
        tmp_dict = {self.columns_name[0]:test_type,
                    self.columns_name[1]:test_text,
                    self.columns_name[2]:page_num,
                    self.columns_name[3]:importance,
                    self.columns_name[4]:right_cnt,
                    self.columns_name[5]:wrong_cnt,
                    self.columns_name[6]:0}
                    #新题的出题次数默认为0

        self.database = self.database.append(tmp_dict,ignore_index=True) 

        #刷新xlsx文件
        self.save_database_to_excel(self.sub_xlsx_file_name +'.xlsx','sheet1')

    def test_num_init(self):

        if self.sub_enum == subject.DVP_PSY:
            self.simple_describ_num = 2
            self.term_expl_num      = 2
            self.describ_num        = 1 
        elif self.sub_enum == subject.EDU_PSY:
            self.simple_describ_num = 0 #教心简答题出题数默认设置为0
            self.term_expl_num      = 1
            self.describ_num        = 1 
        elif self.sub_enum == subject.EXP_PSY:   
            self.simple_describ_num = 2
            self.term_expl_num      = 1
            self.describ_num        = 0 #实验论述题出题数默认设置为0 
        elif self.sub_enum == subject.COMM_PSY: 
            self.simple_describ_num = 1
            self.term_expl_num      = 2
            self.describ_num        = 1 
        elif self.sub_enum == subject.STA_PSY:
            self.simple_describ_num = 1
            self.term_expl_num      = 1
            self.describ_num        = 0 #统计没有论述 
        elif self.sub_enum == subject.SOC_PSY:
            self.simple_describ_num = 2
            self.term_expl_num      = 2
            self.describ_num        = 0 # 社心论述题出题数默认设置为0 
        elif self.sub_enum == subject.MES_PSY:
            self.simple_describ_num = 0 # 测量简答题默认出题数为0
            self.term_expl_num      = 1
            self.describ_num        = 1 # 题库中测量的论述题为1
        else:
            print("Cannot find subject you enter!")
            self.simple_describ_num = 0 # 测量简答题默认出题数为0
            self.term_expl_num      = 0
            self.describ_num        = 0 

    def quiz_setter(self,type:str,num:int,importance:int=0,repeat:bool=False,mark_test:bool=True):
        '''
        功能：出题器 
            type：出题类型,str:[“简答”，“名词解释”,“论述”]
            num：题目数量
            importance：题目最低的重要性，默认为0
            repeat:是否重复，如果设置为True，则可能出现之前出过的题
            mark_test:是否标记出题次数，默认标记
        '''            
        # columns_name = ["题目类型",
        #                 "内容",
        #                 "页码","重要性",
        #                 "答对次数",
        #                 "答错次数",
        #                 "出题次数"]
        if repeat: 
            #出现之前做过题
            pd_tmp = self.database[(self.database["题目类型"] == type) & (self.database['重要性']>=importance)]
        else:
            #只从没做过的题中抽取
            pd_tmp = self.database[(type == self.database["题目类型"]) & (self.database['重要性']>=importance) & (self.database['出题次数'] ==0)]
        # print("----------------------------")
        # print(self.database[(type == self.database["题目类型"])])
        # print("----------------------------")
        if (len(pd_tmp.index) == 0 and num != 0): # 获取df长度
            print("Error: 在",self.subject_name,"题库中没有找到符合要求的",type,"题")
        elif (len(pd_tmp.index) < num):
            print("Error: 在",self.subject_name,"题库中没有找到足够符合要求的",type,"题")
        elif (pd_tmp["重要性"].sum() == 0 ):
            #如果重要性之和为0，就不能使用权重抽样
            pd_tmp = pd_tmp.sample(n=num)
        else:
            #从符合条件的题目中根据重要性随机抽样
            pd_tmp = pd_tmp.sample(n=num,weights=pd_tmp["重要性"])

        if(mark_test and len(pd_tmp.index) >= 0):
            test_sampled_idx_list = pd_tmp.index
            self.database.loc[test_sampled_idx_list,"出题次数"] += 1
            # self.save_database_to_excel(self.sub_xlsx_file_name + '.xlsx','sheet1')            

        return pd_tmp

    def auto_quiz_setter_sim_des(self,mark_test:bool=True):
        '''
           简答题出题器，根据科目类型出题,题目难度默认为0
               mark_test : 出题时是否反标题库，默认打开，测试模式下可关闭
        '''
        pd_sim_des = self.quiz_setter("简答",self.simple_describ_num,self.sim_des_impt,mark_test=mark_test)
        return pd_sim_des

    def auto_quiz_setter_des(self,mark_test:bool=True):
        '''
           论述题出题器，根据科目类型出题,题目难度默认为0
               mark_test : 出题时是否反标题库，默认打开，测试模式下可关闭
        '''
        pd_dscb = self.quiz_setter("论述",self.describ_num,self.des_impt,mark_test=mark_test)
        return pd_dscb

    def auto_quiz_setter_term_expl(self,mark_test:bool=True):
        '''
           名词解释出题器，根据科目类型出题,题目难度默认为0
               mark_test : 出题时是否反标题库，默认打开，测试模式下可关闭
        '''
        pd_term_expl = self.quiz_setter("名词解释",self.term_expl_num,self.term_expl_impt,mark_test=mark_test)
        return pd_term_expl

    def change_wrong_num(self,test_index,change_num):
        '''
            修改本科目下某个题目的答错次数。
            test_index : 数组格式，每个元素为需要修改的题号
            change_num : 数组格式，每个元素为对应题号答错次数的增量
        '''
        if len(test_index) != len(change_num):
            print("输入格式错误，题号和修改量数目不一致")
        else:
            for i in range(len(test_index)):
                self.database.loc[[test_index[i]],["答错次数"]] += change_num[i]

        #刷新xlsx文件
        self.save_database_to_excel(self.sub_xlsx_file_name+'.xlsx','sheet1')
        
    def change_right_num(self,test_index,change_num):
        '''
            修改本科目下某个题目的答对次数。
            test_index : 数组格式，每个元素为需要修改的题号
            change_num : 数组格式，每个元素为对应题号答错次数的增量
        '''
        if len(test_index) != len(change_num):
            print("输入格式错误，题号和修改量数目不一致")
        elif len(test_index) == 0:
            print("输入格式错误")
        else:
            for i in range(len(test_index)):
                self.database.loc[[test_index[i]],["答对次数"]] += change_num[i]     
        #刷新xlsx文件
        self.save_database_to_excel(self.sub_xlsx_file_name+'.xlsx','sheet1')           


##################################################################################################
if __name__ == "__main__":
    # print(subject.DVP_PSY)

    # subjects = [] #初始化空列表
    # i = 0
    # for sub in subject:
    #     subjects.append(subject_db(sub,i))
    #     i += 1

    # for psy_sub in subjects:
    #     psy_sub.parser_subject()


    subject_db_dvp = subject_db(subject.MES_PSY,0)

    subject_db_dvp.get_data_form_excel(subject_db_dvp.sub_xlsx_file_name + '.xlsx')


    print("######################################################")
    print(subject_db_dvp.database)
    print("######################################################")
    # pd_des = subject_db_dvp.auto_quiz_setter_des()
    # print(pd_des)
    # print("######################################################")
    # pd_des = subject_db_dvp.auto_quiz_setter_sim_des()
    # print(pd_des)
    # print("######################################################")
    pd_des = subject_db_dvp.auto_quiz_setter_des(False)
    print(pd_des)
    print("######################################################")
    # print(pd_des.index)
    # subject_db_dvp.change_wrong_num(pd_des.index,[1]*len(pd_des.index))

    # subject_db_dvp.add_new_test('简答','新题目',455,3)
    # print("######################################################")
    # print(subject_db_dvp.database)
    # print("######################################################")
    # subject_db_dvp.parser_subject()
    # subject_db_dvp.save_database_to_excel(r'发展心理学_test.xlsx',"发心")
    # print(subject_db_dvp.database)

