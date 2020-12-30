#from Marc.Docx.DocxThread import downloadThread
import itertools

class DocxRef:

    #docxThread  = downloadThread()

    def check_contain_chinese(self, check_str):  # 判断字符串是否含有中文
        for _char in check_str:
            if '\u4e00' <= _char <= '\u9fa5': return True
        return False

    def merge_author(self, merge_arr, ref_obj):
        '''
        :param merge_arr:
        :param ref_obj:
        :return:
        '''
        for obj in merge_arr:
            if obj['author'] == ref_obj['author']:
                obj['year'] = obj['year'].join(ref_obj['year'])
                obj['ref'] = obj['ref'].join(ref_obj['ref'])

        return merge_arr

    def search_equals_names(self, author, comments_msg, num_arr):
        '''
        :param authors: ['Minshull T A', ' White R S']
        :param ref_list: 所有的参考文献列表
        :return:
        '''
        merge_arr = []
        cursor_num = comments_msg.index(author)
        #print("当前对象的下标是：", cursor_num)
        for author_obj in comments_msg:
            # if cursor_num in num_arr:
            #     continue
            # else:
            #     num_arr.append(cursor_num)

            if len(author['a_names']) <= 2 and len(author_obj['a_names']) <= 2: #作者数小于等于2，必须完全相同
                if author['a_names'] == author_obj['a_names'] :
                    if cursor_num != comments_msg.index(author_obj):
                        if comments_msg.index(author_obj) not in num_arr:
                            merge_arr.append(author_obj)
                            num_arr.append(comments_msg.index(author_obj))
                        #comments_msg.remove(author_obj)
            else:
                if author['a_names'][0] == author_obj['a_names'][0] and len(author_obj['a_names']) >2 and len(author['a_names']) >2:
                    if cursor_num != comments_msg.index(author_obj):
                        if comments_msg.index(author_obj) not in num_arr:
                            merge_arr.append(author_obj)
                            num_arr.append(comments_msg.index(author_obj))
                        #comments_msg.remove(author_obj)
        # print('-------------------------------------------')
        # print(merge_arr)
        # print('-------------------------------------------')
        return merge_arr

    def get_firstname(self, name):
        """
        :param name: 参考文献作者全名
        :return: 返回作者的姓氏
        """
        name_arr = name.split()
        names = []
        for n in name_arr:
            if len(n.strip()) > 1:
                if self.check_contain_chinese(n.strip()):
                    names.append(n)
                else:
                    if n.strip().istitle(): names.append(n)
        return ' '.join(names)


    def strQ2B(self, str):
        """把字符串全角转半角"""
        ss = []
        for s in str:
            rstring = ""
            for uchar in s:
                inside_code = ord(uchar)
                if inside_code == 12288:  # 全角空格直接转换
                    inside_code = 32
                elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
                    inside_code -= 65248
                rstring += chr(inside_code)
            ss.append(rstring)
        return ''.join(ss)

    def create_ref_abbr(self, ref, err_reflist, _comments_mgs):
        """
        :param ref: 一条参考文献
        :return:
        """
        ref_arr = ref.split('.', 2)
        if len(ref_arr) == 3:
            authors = ref_arr[0].strip()
            authors = self.strQ2B(authors)
            authors_arr = authors.split(',')
            year = ref_arr[1].strip()
            print(ref_arr, authors, authors_arr)
            if len(authors_arr) == 1:  # 一位英文作者的情况
                _ref_authors = self.ref_abbr(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            elif len(authors_arr) == 2:  # 两位作者的情况
                _ref_authors = self.ref_abbr2(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            elif len(authors_arr) > 2:  # 多位作者的情况
                _ref_authors = self.ref_abbr3(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            else:
                print('--------------错误，无作者')

        else:
            err_reflist.append(ref)


    def deal_many_authors(self, author, similar_authors):
        similar_authors.insert(0, author)
        author_names = author['a_names']
        year = author['year']
        group_authors_all = []
        comments_msg = []
        for step in range(2, 5):  #len(similar_authors) + 1
            iter1 = itertools.combinations(similar_authors, step)
            group_authors = list(list(t1) for t1 in iter1)
            for au in group_authors:
                group_authors_all.append(au)
        #print(group_authors_all)
        if len(author_names) == 1:
            author_name = author_names[0].strip().split()[0]
            for author1 in  group_authors_all:
                years_str = "、".join([auth["year"]for auth in author1])
                ref = '\r\n'.join([auth["ref"] for auth in author1])
                ref_author = self.ref_abbr(ref, author_names, years_str)
                comments_msg.append(ref_author)
        elif len(author_names) == 2:
            for author1 in  group_authors_all:
                years_str = "、".join([auth["year"]for auth in author1])
                ref = '\r\n'.join([auth["ref"] for auth in author1])
                ref_author = self.ref_abbr2(ref, author_names, years_str)
                comments_msg.append(ref_author)
            #print(f'{author_names}')
        else:
            for author1 in  group_authors_all:
                years_str = "、".join([auth["year"]for auth in author1])
                ref = '\r\n'.join([auth["ref"] for auth in author1])
                ref_author = self.ref_abbr3(ref, author_names, years_str)
                comments_msg.append(ref_author)
        #print(comments_msg)
        return comments_msg

    def deal_authors_merge(self, _comments_mgs, download_proess_signal):
        num_arr = []
        com_msg = []
        i = 1
        for _author in _comments_mgs:
            similar_author = self.search_equals_names(_author, _comments_mgs, num_arr)
            #print(similar_author)
            if len(similar_author) > 0:
                num_arr.append(_comments_mgs.index(_author))
                arr = self.deal_many_authors(_author, similar_author)
                #print('arr:==========',arr)
                for arr_obj in arr:
                    com_msg.append(arr_obj)

            num = int(i / len(_comments_mgs) * 5)
            if num != 100:
                download_proess_signal.emit(num)

            i = i + 1

        return com_msg

    def deal_authors_by_ref(self, ref_list, download_proess_signal):
        '''
        :param ref_list: 根据参考文献列表获取参考文献缩写
        :return:
        '''
        err_refList = []  # 记录错误符号的参考文献
        _comments_mgs = []
        i = 1
        for ref in ref_list:
            self.create_ref_abbr(ref, err_refList, _comments_mgs)

            num = int(i / len(ref_list) * 5)
            if num != 100:
                download_proess_signal.emit(num)
            i = i + 1
        return _comments_mgs

    #@staticmethod
    def ref_abbr(self, ref, authors_arr, year):
        """
        :param ref: 参考文献全文
        :param authors_arr: #['Brown M M'] 一位英文作者 或者 ['陈吉余'] 一位中文作者
        :param year: 参考文献年份
        :return: 
        """
        _ref_authors = {}
        _name = []
        first_name = self.get_firstname(authors_arr[0].strip())
        _name.append(first_name + '（' + year + "）")
        _name.append(first_name + '(' + year + ")")
        _name.append(first_name + '，' + year)
        _name.append(first_name + ',' + year)
        _name_err = []
        _name_err.append(first_name +' {1,}' + '（' + year + "）")
        _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + "）")
        _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + " {1,}）")

        _name_err.append(first_name + ' {1,}' + '(' + year + ")")
        _name_err.append(first_name + ' {1,}' + '( {1,}' + year + ")")
        _name_err.append(first_name + ' {1,}' + '( {1,}' + year + " {1,})")

        _name_err.append(first_name + ' {1,}，' + year)
        _name_err.append(first_name + ' {1,}， {1,}' + year)

        _name_err.append(first_name + ' {1,},' + year)
        _name_err.append(first_name + ' {1,}, {1,}' + year)

        _name1 = []
        _name1.append(first_name + ' ' + '（' + year + "）")
        _name1.append(first_name + ' ' + '（ ' + year + "）")
        _name1.append(first_name + ' ' + '（ ' + year + " ）")

        _name1.append(first_name + ' ' + '(' + year + ")")
        _name1.append(first_name + ' ' + '( ' + year + ")")
        _name1.append(first_name + ' ' + '( ' + year + " )")

        _name1.append(first_name + ' ，' + year)
        _name1.append(first_name + ' ， ' + year)

        _name1.append(first_name + ' ,' + year)
        _name1.append(first_name + ' , ' + year)



        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        _ref_authors['names_err'] = _name_err
        _ref_authors['names1'] = _name1
        return _ref_authors

    # '''
    #     authors_arr : #['Brown M'] 一位中文作者
    #     ref_arr:
    # '''
    #
    # def _ref_abbr_cn(self, ref, authors_arr, year):  #
    #     print(authors_arr)
    #     if '，' in authors_arr[0]:
    #         return self._ref_abbr3_cn(ref, authors_arr[0].split('，'), year)
    #     else:
    #
    #         _ref_authors = {}
    #         _name = []
    #         first_name = authors_arr[0].strip().split(' ')[0].strip()
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + '（' + year + "）")
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + '(' + year + ")")
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + '，' + year)
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + ',' + year)
    #
    #         _name.append(first_name + '（' + year + "）")
    #         _name.append(first_name + '(' + year + ")")
    #         _name.append(first_name + '，' + year)
    #         _name.append(first_name + ',' + year)
    #         _name_err = []
    #         _name_err.append(first_name + ' {1,}' + '（' + year + "）")
    #         _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + "）")
    #         _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + " {1,}）")
    #
    #         _name_err.append(first_name + ' {1,}' + '(' + year + ")")
    #         _name_err.append(first_name + ' {1,}' + '( {1,}' + year + ")")
    #         _name_err.append(first_name + ' {1,}' + '( {1,}' + year + " {1,})")
    #
    #         _name_err.append(first_name + ' {1,}，' + year)
    #         _name_err.append(first_name + ' {1,}， {1,}' + year)
    #
    #         _name_err.append(first_name + ' {1,},' + year)
    #         _name_err.append(first_name + ' {1,}, {1,}' + year)
    #
    #         _ref_authors['ref'] = ref
    #         _ref_authors['names'] = _name
    #         _ref_authors['a_names'] = authors_arr
    #         _ref_authors['year'] = year
    #         _ref_authors['names_err'] = _name_err
    #
    #         #print(_ref_authors)
    #         return _ref_authors

    def ref_abbr2(self, ref, authors_arr, year):
        _ref_authors = {}
        _name = []
        _name_err = []
        _name1 = []
        first_name = self.get_firstname(authors_arr[0].strip())
        sec_name = self.get_firstname(authors_arr[1].strip())
        if not self.check_contain_chinese(ref):  # 纯英文文献
            _name.append(first_name + '和' + sec_name + '（' + year + "）")
            _name.append(first_name + '和' + sec_name + '(' + year + ")")
            _name.append(first_name + ' and ' + sec_name + '，' + year)
            _name.append(first_name + ' and ' + sec_name + ',' + year)
            _name.append(first_name + ' and ' + sec_name + '(' + year + ')')
            _name.append(first_name + ' and ' + sec_name + '（' + year + '）')

            _name_err.append(first_name + ' {1,}和' + sec_name + '（' + year + "）")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '（' + year + "）")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "）")
            #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "） {1,}")

            _name_err.append(first_name + ' {1,}和' + sec_name + '(' + year + ")")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '(' + year + ")")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ")")
            #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ") {1,}")

            _name_err.append(first_name + ' and ' + sec_name + ' {1,}，' + year)
            _name_err.append(first_name + ' and ' + sec_name + ' {1,}， {1,}' + year)

            _name_err.append(first_name + ' and ' + sec_name + ' {1,},' + year)
            _name_err.append(first_name + ' and ' + sec_name + ' {1,}, {1,}' + year)


            _name1.append(first_name + ' 和' + sec_name + '（' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + '（' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + ' （' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + " ）")
            #_name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + "）")
            #_name1.append(first_name + ' 和 ' + sec_name + '（' + year + " ）")

            _name1.append(first_name + ' 和' + sec_name + '(' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + '(' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + ' (' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + " )")

            _name1.append(first_name + ' and ' + sec_name + ' ，' + year)
            _name1.append(first_name + ' and ' + sec_name + ' ， ' + year)

            _name1.append(first_name + ' and ' + sec_name + ' ,' + year)
            _name1.append(first_name + ' and ' + sec_name + ' , ' + year)

        else:
            if not self.check_contain_chinese(authors_arr[0].strip()):
                first_name = self.get_firstname(authors_arr[0].strip())
            else:
                first_name = authors_arr[0].strip()

            if not self.check_contain_chinese(authors_arr[1].strip()):
                sec_name = self.get_firstname(authors_arr[1].strip())
            else:
                sec_name = authors_arr[1].strip()

            _name.append(first_name + '和' + sec_name + '（' + year + "）")
            _name.append(first_name + '和' + sec_name + '(' + year + ")")
            _name.append(first_name + '和' + sec_name + '，' + year)
            _name.append(first_name + '和' + sec_name + ',' + year)
            _name.append(first_name + '、' + sec_name + '，' + year)
            _name.append(first_name + '、' + sec_name + ',' + year)

            _name_err.append(first_name + ' {1,}和' + sec_name + '（' + year + "）")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '（' + year + "）")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "）")
            #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "） {1,}")

            _name_err.append(first_name + ' {1,}和' + sec_name + '(' + year + ")")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '(' + year + ")")
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ")")
            #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ") {1,}")

            _name_err.append(first_name + ' {1,}和' + sec_name + '，' + year)
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '，' + year)
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}，' + year)

            _name_err.append(first_name + ' {1,}和' + sec_name + ',' + year)
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ',' + year)
            _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,},' + year)

            _name_err.append(first_name + ' {1,}、' + sec_name + '，' + year)
            _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + '，' + year)
            _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + ' {1,}，' + year)

            _name_err.append(first_name + ' {1,}、' + sec_name + ',' + year)
            _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + ',' + year)
            _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + ' {1,},' + year)

            _name1.append(first_name + ' 和' + sec_name + '（' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + '（' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + ' （' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + "）")
            _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + " ）")

            _name1.append(first_name + ' 和' + sec_name + '(' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + '(' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + ' (' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + ")")
            _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + " )")

            _name1.append(first_name + ' 和' + sec_name + '，' + year)
            _name1.append(first_name + ' 和 ' + sec_name + '，' + year)
            _name1.append(first_name + ' 和 ' + sec_name + ' ，' + year)
            _name1.append(first_name + ' 和 ' + sec_name + ' ， ' + year)

            _name1.append(first_name + ' 和' + sec_name + ',' + year)
            _name1.append(first_name + ' 和 ' + sec_name + ',' + year)
            _name1.append(first_name + ' 和 ' + sec_name + ' ,' + year)
            _name1.append(first_name + ' 和 ' + sec_name + ' , ' + year)

            _name1.append(first_name + ' 、' + sec_name + '，' + year)
            _name1.append(first_name + ' 、 ' + sec_name + '，' + year)
            _name1.append(first_name + ' 、 ' + sec_name + ' ，' + year)
            _name1.append(first_name + ' 、 ' + sec_name + ' ， ' + year)

            _name1.append(first_name + ' 、' + sec_name + ',' + year)
            _name1.append(first_name + ' 、 ' + sec_name + ',' + year)
            _name1.append(first_name + ' 、 ' + sec_name + ' ,' + year)
            _name1.append(first_name + ' 、 ' + sec_name + ' , ' + year)
        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        _ref_authors['names_err'] = _name_err
        _ref_authors['names1'] = _name1
        #print(_ref_authors)
        return _ref_authors



    # def _ref_abbr2_cn(self, ref, authors_arr, year):
    #     '''
    #         authors_arr :  ['高原', ' 郑斯华'] 两位中文作者
    #         ref_arr:
    #     '''
    #     _ref_authors = {}
    #     _name = []
    #     first_name = ''
    #     sec_name = ''
    #     if not self.check_contain_chinese(authors_arr[0].strip()):
    #         first_name = authors_arr[0].strip().split(' ')[0].strip()
    #     else:
    #         first_name = authors_arr[0].strip()
    #
    #     if not self.check_contain_chinese(authors_arr[1].strip()):
    #         sec_name = authors_arr[1].strip().split(' ')[0].strip()
    #     else:
    #         sec_name = authors_arr[1].strip()
    #
    #     _name.append(first_name + '和' + sec_name + '（' + year + "）")
    #     _name.append(first_name + '和' + sec_name + '(' + year + ")")
    #     _name.append(first_name + '和' + sec_name + '，' + year)
    #     _name.append(first_name + '和' + sec_name + ',' + year)
    #     _name.append(first_name + '、' + sec_name + '，' + year)
    #     _name.append(first_name + '、' + sec_name + ',' + year)
    #
    #     _ref_authors['ref'] = ref
    #     _ref_authors['names'] = _name
    #     _ref_authors['a_names'] = authors_arr
    #     _ref_authors['year'] = year
    #
    #     #print(_ref_authors)
    #     return _ref_authors

    '''
        authors_arr : ['Minshull T A', ' Muller M R', ' White R S'], ['Minshull T A', ' Muller M R', ' Robinson C J', ' et al'] 多位作者  
        ref_arr: 
    '''

    def ref_abbr3(self, ref, authors_arr, year):  #
        _ref_authors = {}
        _name = []
        _name_err = []
        _name1 = []
        first_name = self.get_firstname(authors_arr[0].strip())
        if not self.check_contain_chinese(ref):
            _name.append(first_name + '等' + '（' + year + "）")
            _name.append(first_name + '等' + '(' + year + ")")
            _name.append(first_name + ' et al.' + '，' + year)
            _name.append(first_name + ' et al.' + ',' + year)

            _name_err.append(first_name + ' {1,}等' + '（' + year + "）")
            _name_err.append(first_name + ' {1,}等 {1,}' + '（' + year + "）")
            _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + "）")
            _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")

            _name_err.append(first_name + ' {1,}等' + '(' + year + ")")
            _name_err.append(first_name + ' {1,}等 {1,}' + '(' + year + ")")
            _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + ")")
            _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + " {1,})")

            _name_err.append(first_name + '  {1,}et al.' + '，' + year)
            _name_err.append(first_name + '  {1,}et al {1,}.' + '，' + year)
            _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + '，' + year)
            _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + '， {1,}' + year)

            _name_err.append(first_name + '  {1,}et al.' + ',' + year)
            _name_err.append(first_name + '  {1,}et al {1,}.' + ',' + year)
            _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + ',' + year)
            _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + ', {1,}' + year)

            _name1.append(first_name + ' 等' + '（' + year + "）")
            _name1.append(first_name + ' 等 ' + '（' + year + "）")
            _name1.append(first_name + ' 等 ' + '（ ' + year + "）")
            _name1.append(first_name + ' 等 ' + '（ ' + year + " ）")

            _name1.append(first_name + ' 等' + '(' + year + ")")
            _name1.append(first_name + ' 等 ' + '(' + year + ")")
            _name1.append(first_name + ' 等 ' + '( ' + year + ")")
            _name1.append(first_name + ' 等 ' + '( ' + year + " )")

            _name1.append(first_name + '  et al.' + '，' + year)
            _name1.append(first_name + '  et al .' + '，' + year)
            _name1.append(first_name + '  et al . ' + '，' + year)
            _name1.append(first_name + '  et al . ' + '， ' + year)

            _name1.append(first_name + '  et al.' + ',' + year)
            _name1.append(first_name + '  et al .' + ',' + year)
            _name1.append(first_name + '  et al . ' + ',' + year)
            _name1.append(first_name + '  et al . ' + ', ' + year)

        else:
            _name.append(first_name + '等' + '（' + year + "）")
            _name.append(first_name + '等' + '(' + year + ")")
            _name.append(first_name + '等' + '，' + year)
            _name.append(first_name + '等' + ',' + year)

            _name_err.append(first_name + ' {1,}等' + '（' + year + "）")
            _name_err.append(first_name + ' {1,}等 {1,}' + '（' + year + "）")
            _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + "）")
            _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")

            _name_err.append(first_name + ' {1,}等' + '(' + year + ")")
            _name_err.append(first_name + ' {1,}等 {1,}' + '(' + year + ")")
            _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + ")")
            _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + " {1,})")

            _name_err.append(first_name + ' {1,}等' + '，' + year)
            _name_err.append(first_name + ' {1,}等 {1,}' + '，' + year)
            _name_err.append(first_name + ' {1,}等 {1,}' + '， {1,}' + year)
            #_name_err.append(first_name + ' {1,}等 {1,}' + '， {1,}' + year)

            _name_err.append(first_name + ' {1,}等' + ',' + year)
            _name_err.append(first_name + ' {1,}等 {1,}' + ',' + year)
            _name_err.append(first_name + ' {1,}等 {1,}' + ', {1,}' + year)
            #_name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")
            _name1.append(first_name + ' 等' + '（' + year + "）")
            _name1.append(first_name + ' 等 ' + '（' + year + "）")
            _name1.append(first_name + ' 等 ' + '（ ' + year + "）")
            _name1.append(first_name + ' 等 ' + '（ ' + year + " ）")

            _name1.append(first_name + ' 等' + '(' + year + ")")
            _name1.append(first_name + ' 等 ' + '(' + year + ")")
            _name1.append(first_name + ' 等 ' + '( ' + year + ")")
            _name1.append(first_name + ' 等 ' + '( ' + year + " )")

            _name1.append(first_name + ' 等' + '，' + year)
            _name1.append(first_name + ' 等 ' + '，' + year)
            _name1.append(first_name + ' 等 ' + '， ' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + '， {1,}' + year)

            _name1.append(first_name + ' 等' + ',' + year)
            _name1.append(first_name + ' 等 ' + ',' + year)
            _name1.append(first_name + ' 等 ' + ', ' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")


        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        _ref_authors['names_err'] = _name_err
        _ref_authors['names1'] = _name1
        return _ref_authors



    # def _ref_abbr3_cn(self, ref, authors_arr, year):  #
    #     """
    #         authors_arr : ['卫小冬', ' 赵明辉', ' 阮爱国等'] 多位中文作者
    #         ref_arr:
    #     """
    #     _ref_authors = {}
    #     _name = []
    #     _name_err = []
    #
    #
    #     _name.append(authors_arr[0].strip().split(' ')[0].strip() + '等' + '（' + year + "）")
    #     _name.append(authors_arr[0].strip().split(' ')[0].strip() + '等' + '(' + year + ")")
    #     _name.append(authors_arr[0].strip().split(' ')[0].strip() + '等' + '，' + year)
    #     _name.append(authors_arr[0].strip().split(' ')[0].strip() + '等' + ',' + year)
    #
    #     _ref_authors['ref'] = ref
    #     _ref_authors['names'] = _name
    #     _ref_authors['a_names'] = authors_arr
    #     _ref_authors['year'] = year
    #     _ref_authors['names_err'] = _name_err
    #     #print(_ref_authors)
    #     return _ref_authors


if __name__ == '__main__':
    authors_arr = ['Van Maren D S', 'Oost A P', 'Wang Z B','Vos P C','LeBlond P H','G']
    for name in authors_arr:
        name_arr = name.split()
        print(name_arr)
        names = []
        for n in name_arr:
            if len(n.strip())!=1:names.append(n)
        print(names)

    first_name = authors_arr[0].strip().split()[0].strip()
    print()

    print('国家'.istitle())
    print('Maren'.istitle())
    print('MMen'.istitle())




