import docx
import comtypes.client
import win32com.client as cc
import sys, os, time, re, subprocess, datetime

def generate_cv(filename,date,to,rank,company,company_loc,position):
    document = docx.Document("template/"+filename)

    
    keywords = ['<date>','<to>','<rank>','<company>','<company_loc>','<position>']
    values = [date,to,rank,company,company_loc,position]

    for para in document.paragraphs:
        try:
            print(para.text)
            temp = str(para.text.encode('ascii', 'ignore'))

            para_in_list = temp.split(' ')
            
            for key in keywords:
                print('Checking for: '+key)
                while key in para_in_list:
                    print(key+' found!')
                    print('Found in {}'.format(para_in_list))
                    time.sleep(0.5)
                    key_index = para_in_list.index(key)
                    print('From {0}'.format(para_in_list))
                    para_in_list[key_index] = values[keywords.index(key)]
                    print('To {0}'.format(para_in_list))
                else:
                    pass
                    #print(key+' Not Found!')
                #print('Working...')
            
            para_in_list = " ".join(str(x) for x in para_in_list)
            para.text = para_in_list
            para.style.font.name = 'Arial'
        except Exception as e:
            print('\n{}\n'.format(e))

    final_filename = '{0}_for_{1}_{2}.docx'.format(filename.replace('.docx',''),company,position)
    path = '{0}_{1}/'.format(company,position)
    
    if not os.path.exists('results/'+path):
        os.makedirs('results/'+path)
    ff = 'results/'+path+final_filename

    print('Saving new file...')
    document.save(ff)
    time.sleep(1)

def main_screen():
    subprocess.call('cls',shell=True)
    print('*************************************************************************************************************************************')
    print('* Resume Editor Script by Mark P                                                                                                    *')
    print('* Type:                                                                                                                             *')
    print('* python generate_v2.py "[Hiring Manger]" "[HR position]" "[company_name]" "[company,location]" "[position]"                        *')
    print('*                                                                                                                                   *')
    print('* Example:                                                                                                                          *')
    print('* python generate_v2.py "Mark Prado" "HR Manager" "Sheridan_College" "7899 McLaughlin Rd,Brampton,ON L6Y 5H9" "Software Engineer"   *')
    print('*                                                                                                                                   *')
    print('*************************************************************************************************************************************')
    print ('\n')  

if __name__ == "__main__":
    #command line syntax:
    #python generate.py <company> <company_loc> <apply_pos>
    main_screen()
    
    d = datetime.datetime.now()
    date = str(d.strftime("%B %d, %Y"))
    print ("Date: {}".format(date))

    to = sys.argv[1]
    print ("To: {}".format(to))

    rank = sys.argv[2]
    print ("Rank: {}".format(rank))

    company = sys.argv[3]
    print ("Company: {}".format(company))

    company_loc = sys.argv[4]
    company_loc = company_loc.replace(',',',\n')
    print ("Company Location: {}".format(company_loc))

    position = sys.argv[5]
    print ("Position Applied: {}".format(position))

    for doc in os.listdir("template/"):
        if '~' not in doc:
            generate_cv(doc,date,to,rank,company,company_loc,position)
    
    main_screen()
    print('DONE!')

    