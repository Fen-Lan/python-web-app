import json

class Element():

    def __init__(self, ele_file): 
        with open(ele_file, 'r', encoding='UTF-8') as file:
            self.eles_list = json.load(file)

    def get_xpath(self, name:str, replace_text:str = None) -> str:
        '''
        获取指定元素的xpath值
        name: 元素的名称
        replace_text: 待替换的值,默认为None
        '''
        for group in self.eles_list:
            controls = group['controls']
            for ctrl in controls:
                n = ctrl['name']
                xpath = ctrl['xpath']
                if n == name:
                    if not replace_text:
                        return xpath
                    
        else:
            return 'None'

    def get_iframes(self, name):
        iframes = []
        for group in self.eles_list:
            controls = group['controls']
            for ctrl in controls:
                n = ctrl['name']
                iframes = ctrl['iframes']
                if n == name:
                    return iframes
        else:
            return iframes