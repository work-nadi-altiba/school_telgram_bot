from fillpdf import fillpdfs
import pandas as pd
import math
import re

# print(fillpdfs.get_form_fields('students_data.pdf'))

forms = {'text_1xbky': '', 'text_2xoaz': '', 'text_3isqe': '', 'text_4pdsa': '', 'text_5adol': '', 'text_6epkt': '', 'text_7fyef': '', 'text_8oift': '', 'text_9nzdn': '', 'text_10xtuq': '', 'text_11nf': '', 'text_12bwlh': '', 'text_13dwdc': '', 'text_14sqwe': '', 'text_15nwhp': '', 'text_16kmrr': '', 'text_17inqk': '', 'text_18xgsi': '', 'text_19dfuw': '', 'text_20bsry': '', 'text_21tefr': '', 'text_22eymr': '', 'text_23qdwx': '', 'text_24pfbm': '', 'text_25goqm': '', 'text_26qvlp': '', 'text_27lmsw': '', 'text_28nfvb': '', 'text_29vcyr': '', 'text_30vbhd': '', 'text_31ohvz': '', 'text_32sfnh': '', 'text_33roal': '', 'text_34fvu': '', 'text_35cpgi': '', 'text_36gzhj': '', 'text_37fspf': '', 'text_38ejge': '', 'text_39pnue': '', 'text_40tiaq': '', 'text_41oiww': '', 'text_42ydnm': '', 'text_43hnqy': '', 'text_44xmih': '', 'text_45fihc': '', 'text_46fmjt': '', 'text_47ygjw': '', 'text_48evif': '', 'text_49mkvb': '', 'text_50aotk': '', 'text_51aixn': '', 'text_52qoab': '', 'text_53nggj': '', 'text_54nwva': '', 'text_55evsv': '', 'text_56xjst': '', 'text_57yxxt': '', 'text_58ngcy': '', 'text_59jrpj': '', 'text_60nnyq': '', 'text_61lub': '', 'text_62miiw': '', 'text_63elas': '', 'text_64tgzz': '', 'text_65bfb': '', 'text_66jwmp': '', 'text_67guxn': '', 'text_68ovqy': '', 'text_69owlg': '', 'text_70izba': '', 'text_71jndc': '', 'text_72asto': '', 'text_73xonv': '', 'text_74hkdz': '', 'text_75lvhd': '', 'text_76qksg': '', 'text_77pkxg': '', 'text_78p': '', 'text_79pxn': '', 'text_80xt': '', 'text_81ilze': '', 'text_82sjam': '', 'text_83ncpn': '', 'text_84ezib': '', 'text_85ayjn': '', 'text_86kjpy': '', 'text_87desb': '', 'text_88wire': '', 'text_89ijsu': '', 'text_90mumd': '', 'text_91xhpg': '', 'text_92unvd': '', 'text_93skuc': '', 'text_94vb': '', 'text_95bzpf': '', 'text_96wbm': '', 'text_97lasa': '', 'text_98zjtg': '', 'text_99rrcs': '', 'text_100xazy': '', 'text_101olev': '', 'text_102zgum': '', 'text_103hezr': '', 'text_104mzrn': '', 'text_105rayr': '', 'text_106bhfz': '', 'text_107xozq': '', 'text_108axzf': '', 'text_109pejm': '', 'text_110bxqu': '', 'text_111fksx': '', 'text_112yxdy': '', 'text_113pqo': '', 'text_114vgix': '', 'text_115gbwa': '', 'text_116ppux': '', 'text_117hbip': '', 'text_118kfmn': '', 'text_119ybsl': '', 'text_120dyzj': '', 'text_121bxbj': '', 'text_122tltm': '', 'text_123aogt': '', 'text_124jeax': '', 'text_125hymd': '', 'text_126vbj': '', 'text_127dvig': '', 'text_128npgo': '', 'text_129fpma': '', 'text_130hmsg': '', 'text_131dxp': '', 'text_132ytqu': '', 'text_133wzoh': '', 'text_134uajp': '', 'text_135ztsh': '', 'text_136fb': '', 'text_137pbtx': '', 'text_138sadd': '', 'text_139cplq': '', 'text_140puvt': '', 'text_141fbol': '', 'text_142ptr': '', 'text_143tcbm': '', 'text_144qgjt': '', 'text_145qlsh': '', 'text_146mblj': '', 'text_147mhpr': '', 'text_148jtrr': '', 'text_149qdqj': '', 'text_150kzcq': '', 'text_151fanp': '', 'text_152byaw': '', 'text_153mhkc': '', 'text_154vgid': '', 'text_155wwih': '', 'text_156ynwi': '', 'text_157xvao': '', 'text_158xri': '', 'text_159yzzv': '', 'text_160miqc': '', 'text_161zldg': '', 'text_162ppvu': '', 'text_163mrwm': '', 'text_164fxzl': '', 'text_165jfny': '', 'text_166eupw': '', 'text_167yubk': '', 'text_168bfgw': '', 'text_169mpgr': '', 'text_170axho': '', 'text_171nwxe': '', 'text_172dfje': '', 'text_173qydm': '', 'text_174ycgd': '', 'text_175yzav': '', 'text_176abtu': '', 'text_177qexu': '', 'text_178arkg': '', 'text_179yjdb': '', 'text_180oaet': '', 'text_181ldoj': '', 'text_182ivcw': '', 'text_183muvo': '', 'text_184stqv': '', 'text_185xoh': '', 'text_186zeky': '', 'text_187uvbb': '', 'text_188wgem': '', 'text_189ljla': '', 'text_190ijgh': '', 'text_191cqor': '', 'text_192oyie': '', 'text_193owdp': '', 'text_194qgve': '', 'text_195qkss': '', 'text_196jjzr': '', 'text_197wzgq': '', 'text_198fuqg': '', 'text_199defl': '', 'text_200biuj': '', 'text_201': '', 'text_202yaeg': '', 'text_203jtlh': '', 'text_204nuea': '', 'text_205tvrv': '', 'text_206arhh': '', 'text_207bfnl': '', 'text_208drup': '', 'text_209ikph': '', 'text_210lyxb': '', 'text_211kvmj': '', 'text_212zkjh': '', 'text_213imno': '', 'text_214qgln': '', 'text_215brvj': '', 'text_216qawg': '', 'text_217ys': '', 'text_218qpuq': '', 'text_219fsaf': '', 'text_220skum': '', 'text_221bdd': '', 'text_222aqf': '', 'text_223ucmb': '', 'text_224vbgn': '', 'text_225ivrf': '', 'text_226elfq': '', 'text_227ralv': '', 'text_228zbqt': '', 'text_229tvrz': '', 'text_230dlos': '', 'text_231mviy': '', 'text_232lkxb': '', 'text_233qoip': '', 'text_234hzqp': ''}

lis = ['text_1xbky', 'text_2xoaz', 'text_3isqe', 'text_4pdsa', 'text_5adol', 'text_6epkt', 'text_7fyef', 'text_8oift', 'text_9nzdn', 'text_10xtuq', 'text_11nf', 'text_12bwlh', 'text_13dwdc', 'text_14sqwe', 'text_15nwhp', 'text_16kmrr', 'text_17inqk', 'text_18xgsi', 'text_19dfuw', 'text_20bsry', 'text_21tefr', 'text_22eymr', 'text_23qdwx', 'text_24pfbm', 'text_25goqm', 'text_26qvlp', 'text_27lmsw', 'text_28nfvb', 'text_29vcyr', 'text_30vbhd', 'text_31ohvz', 'text_32sfnh', 'text_33roal', 'text_34fvu', 'text_35cpgi', 'text_36gzhj', 'text_37fspf', 'text_38ejge', 'text_39pnue', 'text_40tiaq', 'text_41oiww', 'text_42ydnm', 'text_43hnqy', 'text_44xmih', 'text_45fihc', 'text_46fmjt', 'text_47ygjw', 'text_48evif', 'text_49mkvb', 'text_50aotk', 'text_51aixn', 'text_52qoab', 'text_53nggj', 'text_54nwva', 'text_55evsv', 'text_56xjst', 'text_57yxxt', 'text_58ngcy', 'text_59jrpj', 'text_60nnyq', 'text_61lub', 'text_62miiw', 'text_63elas', 'text_64tgzz', 'text_65bfb', 'text_66jwmp', 'text_67guxn', 'text_68ovqy', 'text_69owlg', 'text_70izba', 'text_71jndc', 'text_72asto', 'text_73xonv', 'text_74hkdz', 'text_75lvhd', 'text_76qksg', 'text_77pkxg', 'text_78p', 'text_79pxn', 'text_80xt', 'text_81ilze', 'text_82sjam', 'text_83ncpn', 'text_84ezib', 'text_85ayjn', 'text_86kjpy', 'text_87desb', 'text_88wire', 'text_89ijsu', 'text_90mumd', 'text_91xhpg', 'text_92unvd', 'text_93skuc', 'text_94vb', 'text_95bzpf', 'text_96wbm', 'text_97lasa', 'text_98zjtg', 'text_99rrcs', 'text_100xazy', 'text_101olev', 'text_102zgum', 'text_103hezr', 'text_104mzrn', 'text_105rayr', 'text_106bhfz', 'text_107xozq', 'text_108axzf', 'text_109pejm', 'text_110bxqu', 'text_111fksx', 'text_112yxdy', 'text_113pqo', 'text_114vgix', 'text_115gbwa', 'text_116ppux', 'text_117hbip', 'text_118kfmn', 'text_119ybsl', 'text_120dyzj', 'text_121bxbj', 'text_122tltm', 'text_123aogt', 'text_124jeax', 'text_125hymd', 'text_126vbj', 'text_127dvig', 'text_128npgo', 'text_129fpma', 'text_130hmsg', 'text_131dxp', 'text_132ytqu', 'text_133wzoh', 'text_134uajp', 'text_135ztsh', 'text_136fb', 'text_137pbtx', 'text_138sadd', 'text_139cplq', 'text_140puvt', 'text_141fbol', 'text_142ptr', 'text_143tcbm', 'text_144qgjt', 'text_145qlsh', 'text_146mblj', 'text_147mhpr', 'text_148jtrr', 'text_149qdqj', 'text_150kzcq', 'text_151fanp', 'text_152byaw', 'text_153mhkc', 'text_154vgid', 'text_155wwih', 'text_156ynwi', 'text_157xvao', 'text_158xri', 'text_159yzzv', 'text_160miqc', 'text_161zldg', 'text_162ppvu', 'text_163mrwm', 'text_164fxzl', 'text_165jfny', 'text_166eupw', 'text_167yubk', 'text_168bfgw', 'text_169mpgr', 'text_170axho', 'text_171nwxe', 'text_172dfje', 'text_173qydm', 'text_174ycgd', 'text_175yzav', 'text_176abtu', 'text_177qexu', 'text_178arkg', 'text_179yjdb', 'text_180oaet', 'text_181ldoj', 'text_182ivcw', 'text_183muvo', 'text_184stqv', 'text_185xoh', 'text_186zeky', 'text_187uvbb', 'text_188wgem', 'text_189ljla', 'text_190ijgh', 'text_191cqor', 'text_192oyie', 'text_193owdp', 'text_194qgve', 'text_195qkss', 'text_196jjzr', 'text_197wzgq', 'text_198fuqg', 'text_199defl', 'text_200biuj', 'text_201', 'text_202yaeg', 'text_203jtlh', 'text_204nuea', 'text_205tvrv', 'text_206arhh', 'text_207bfnl', 'text_208drup', 'text_209ikph', 'text_210lyxb', 'text_211kvmj', 'text_212zkjh', 'text_213imno', 'text_214qgln', 'text_215brvj', 'text_216qawg', 'text_217ys', 'text_218qpuq', 'text_219fsaf', 'text_220skum', 'text_221bdd', 'text_222aqf', 'text_223ucmb', 'text_224vbgn', 'text_225ivrf', 'text_226elfq', 'text_227ralv', 'text_228zbqt', 'text_229tvrz', 'text_230dlos', 'text_231mviy', 'text_232lkxb', 'text_233qoip', 'text_234hzqp']
 
df = pd.read_csv("out2.csv")
# df = df.dropna().reset_index(drop=True)

# print('text_1xbky' in lis)

def find_form(number, lis=lis):
    filtered_value = list(filter(lambda v: re.match(f'text_{number}' +'[a-z]{0,4}', v), lis))
    # print(filtered_value[0])
    # input('press anything to continue')
    # return filtered_value[0]
    if filtered_value[0] == 'nan':
        return ''
    else:
        return filtered_value[0]
# print(find_form(1))

const = 0
dataframe_pointer = 0
for dataframe in range( 1 , len(df['SN'].values )):
    for row in range(1 , 19):
        # breakpoint()
        t1 = find_form(row +const)
        const += 18
        t2 = find_form(row +const)
        const += 18
        t3 = find_form(row +const )
        const += 18
        t4 = find_form(row +const )
        const += 18
        t5 = find_form(row +const )
        const += 18
        t6 = find_form(row +const )
        const += 18
        t7 = find_form(row +const )
        const += 18
        t8 = find_form(row +const )
        const += 18
        t9 = find_form(row +const )
        const += 18
        t10 = find_form(row +const )
        const += 18
        t11 = find_form(row +const )
        const += 18
        t12 = find_form(row +const )
        const += 18
        t13 = find_form(row +const )
        # breakpoint()
        print(t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13) 
        try : 
            full_name = df['SN'].values[dataframe_pointer].split(' ')
            forms[f'{t1}']  = int(float(df['emis_id'].values[dataframe_pointer]))
            forms[f'{t2}']  = ''
            forms[f'{t3}']  = df['name1'].values[dataframe_pointer]
            forms[f'{t4}']  = df['name2'].values[dataframe_pointer]
            forms[f'{t5}']  = df['name3'].values[dataframe_pointer]
            forms[f'{t6}']  = df['name4'].values[dataframe_pointer]
            forms[f'{t7}']  = df['birth1'].values[dataframe_pointer] 
            forms[f'{t8}']  = df['birth_date'].values[dataframe_pointer] 
            forms[f'{t9}']  = df['nationality'].values[dataframe_pointer] 
            forms[f'{t10}']  = df['gender'].values[dataframe_pointer] 
            forms[f'{t11}']  = df['resedent1'].values[dataframe_pointer] 
            forms[f'{t12}']  = df['resedent2'].values[dataframe_pointer] 
            forms[f'{t13}']  = df['resedent3'].values[dataframe_pointer] 
        except IndexError:
            break
        # breakpoint( )
        print(row-1 , const)
        dataframe_pointer += 1
        # breakpoint()
        const = 0
        # input('press anything to continue')
        # forms.pop(0)
    fillpdfs.write_fillable_pdf('students_data.pdf', f'filled1.pdf', forms, flatten=False)
    forms = {'text_1xbky': '', 'text_2xoaz': '', 'text_3isqe': '', 'text_4pdsa': '', 'text_5adol': '', 'text_6epkt': '', 'text_7fyef': '', 'text_8oift': '', 'text_9nzdn': '', 'text_10xtuq': '', 'text_11nf': '', 'text_12bwlh': '', 'text_13dwdc': '', 'text_14sqwe': '', 'text_15nwhp': '', 'text_16kmrr': '', 'text_17inqk': '', 'text_18xgsi': '', 'text_19dfuw': '', 'text_20bsry': '', 'text_21tefr': '', 'text_22eymr': '', 'text_23qdwx': '', 'text_24pfbm': '', 'text_25goqm': '', 'text_26qvlp': '', 'text_27lmsw': '', 'text_28nfvb': '', 'text_29vcyr': '', 'text_30vbhd': '', 'text_31ohvz': '', 'text_32sfnh': '', 'text_33roal': '', 'text_34fvu': '', 'text_35cpgi': '', 'text_36gzhj': '', 'text_37fspf': '', 'text_38ejge': '', 'text_39pnue': '', 'text_40tiaq': '', 'text_41oiww': '', 'text_42ydnm': '', 'text_43hnqy': '', 'text_44xmih': '', 'text_45fihc': '', 'text_46fmjt': '', 'text_47ygjw': '', 'text_48evif': '', 'text_49mkvb': '', 'text_50aotk': '', 'text_51aixn': '', 'text_52qoab': '', 'text_53nggj': '', 'text_54nwva': '', 'text_55evsv': '', 'text_56xjst': '', 'text_57yxxt': '', 'text_58ngcy': '', 'text_59jrpj': '', 'text_60nnyq': '', 'text_61lub': '', 'text_62miiw': '', 'text_63elas': '', 'text_64tgzz': '', 'text_65bfb': '', 'text_66jwmp': '', 'text_67guxn': '', 'text_68ovqy': '', 'text_69owlg': '', 'text_70izba': '', 'text_71jndc': '', 'text_72asto': '', 'text_73xonv': '', 'text_74hkdz': '', 'text_75lvhd': '', 'text_76qksg': '', 'text_77pkxg': '', 'text_78p': '', 'text_79pxn': '', 'text_80xt': '', 'text_81ilze': '', 'text_82sjam': '', 'text_83ncpn': '', 'text_84ezib': '', 'text_85ayjn': '', 'text_86kjpy': '', 'text_87desb': '', 'text_88wire': '', 'text_89ijsu': '', 'text_90mumd': '', 'text_91xhpg': '', 'text_92unvd': '', 'text_93skuc': '', 'text_94vb': '', 'text_95bzpf': '', 'text_96wbm': '', 'text_97lasa': '', 'text_98zjtg': '', 'text_99rrcs': '', 'text_100xazy': '', 'text_101olev': '', 'text_102zgum': '', 'text_103hezr': '', 'text_104mzrn': '', 'text_105rayr': '', 'text_106bhfz': '', 'text_107xozq': '', 'text_108axzf': '', 'text_109pejm': '', 'text_110bxqu': '', 'text_111fksx': '', 'text_112yxdy': '', 'text_113pqo': '', 'text_114vgix': '', 'text_115gbwa': '', 'text_116ppux': '', 'text_117hbip': '', 'text_118kfmn': '', 'text_119ybsl': '', 'text_120dyzj': '', 'text_121bxbj': '', 'text_122tltm': '', 'text_123aogt': '', 'text_124jeax': '', 'text_125hymd': '', 'text_126vbj': '', 'text_127dvig': '', 'text_128npgo': '', 'text_129fpma': '', 'text_130hmsg': '', 'text_131dxp': '', 'text_132ytqu': '', 'text_133wzoh': '', 'text_134uajp': '', 'text_135ztsh': '', 'text_136fb': '', 'text_137pbtx': '', 'text_138sadd': '', 'text_139cplq': '', 'text_140puvt': '', 'text_141fbol': '', 'text_142ptr': '', 'text_143tcbm': '', 'text_144qgjt': '', 'text_145qlsh': '', 'text_146mblj': '', 'text_147mhpr': '', 'text_148jtrr': '', 'text_149qdqj': '', 'text_150kzcq': '', 'text_151fanp': '', 'text_152byaw': '', 'text_153mhkc': '', 'text_154vgid': '', 'text_155wwih': '', 'text_156ynwi': '', 'text_157xvao': '', 'text_158xri': '', 'text_159yzzv': '', 'text_160miqc': '', 'text_161zldg': '', 'text_162ppvu': '', 'text_163mrwm': '', 'text_164fxzl': '', 'text_165jfny': '', 'text_166eupw': '', 'text_167yubk': '', 'text_168bfgw': '', 'text_169mpgr': '', 'text_170axho': '', 'text_171nwxe': '', 'text_172dfje': '', 'text_173qydm': '', 'text_174ycgd': '', 'text_175yzav': '', 'text_176abtu': '', 'text_177qexu': '', 'text_178arkg': '', 'text_179yjdb': '', 'text_180oaet': '', 'text_181ldoj': '', 'text_182ivcw': '', 'text_183muvo': '', 'text_184stqv': '', 'text_185xoh': '', 'text_186zeky': '', 'text_187uvbb': '', 'text_188wgem': '', 'text_189ljla': '', 'text_190ijgh': '', 'text_191cqor': '', 'text_192oyie': '', 'text_193owdp': '', 'text_194qgve': '', 'text_195qkss': '', 'text_196jjzr': '', 'text_197wzgq': '', 'text_198fuqg': '', 'text_199defl': '', 'text_200biuj': '', 'text_201': '', 'text_202yaeg': '', 'text_203jtlh': '', 'text_204nuea': '', 'text_205tvrv': '', 'text_206arhh': '', 'text_207bfnl': '', 'text_208drup': '', 'text_209ikph': '', 'text_210lyxb': '', 'text_211kvmj': '', 'text_212zkjh': '', 'text_213imno': '', 'text_214qgln': '', 'text_215brvj': '', 'text_216qawg': '', 'text_217ys': '', 'text_218qpuq': '', 'text_219fsaf': '', 'text_220skum': '', 'text_221bdd': '', 'text_222aqf': '', 'text_223ucmb': '', 'text_224vbgn': '', 'text_225ivrf': '', 'text_226elfq': '', 'text_227ralv': '', 'text_228zbqt': '', 'text_229tvrz': '', 'text_230dlos': '', 'text_231mviy': '', 'text_232lkxb': '', 'text_233qoip': '', 'text_234hzqp': ''}
    breakpoint()
    # const = 0
    
        