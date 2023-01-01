import pandas as pd

def json_to_csv():
    df = pd.read_json (r'data3.json')
    df.to_csv (r'output.csv', index = None)

def sort_csv(column , sort=None   ):
    df = pd.read_csv("out2.csv")
    # result = df[df[f'{column}']==f'{sort}']
    sorted_df = df.sort_values(by=[column], ascending=True)
    return sorted_df
    # return result.sort_values(["SN"] , ascending=[True]).reset_index(drop=True)
    # names = result['اسم الطالب'].tolist()
    # return names

def get_names(classname):
    df = pd.read_csv("out2.csv")
    result = df[df['الصف و الشعبة'] == str(classname)]
    names = result['اسم الطالب'].tolist()
    return names

if __name__ == '__main__':
    sort_csv('SN')
    # sort_csv('SN' , 'الصف السادس-د')
    
    # pg = sort_csv('SC' , 'الصف السادس-د')
    # print(pg)
    # json_to_csv()
    # df = sort_csv('SC' , 'الصف السادس-د').fillna('')
    # df.to_csv('out2.csv' , index=False ,encoding='utf-8' )  

    # print(df['name2'].values[0])
    
    # .sort_values(["Age"], axis=0, ascending=[False], inplace=True)
    # ['الصف السادس-أ', 'الصف السادس-ب', 'الصف السادس-ج', 'الصف السادس-د', 'الصف السادس-ر', 'الصف السادس-ز', 'الصف السادس-ع', 'الصف السادس-هـ', 'الصف السادس-و', 'الصف السادس-ي']
