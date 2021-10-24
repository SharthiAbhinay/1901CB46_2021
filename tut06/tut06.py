import os
import re



def rename_Breaking_bad(webseries_name,season_padding,episode_padding):
    folder1 = "correct_srt"
 
    if not os.path.exists(folder1): 
        os.makedirs(folder1)
    season_pad = int(season_padding)
    episode_pad = int(episode_padding)
    basepath = "wrong_srt"
    destiny=os.path.join(os.getcwd(),"correct_srt",webseries_name)
    path=os.path.join(os.getcwd(),"tut06\wrong_srt",webseries_name)
    dest=os.path.join(os.getcwd(),"correct_srt")
    folder_path = os.path.join(basepath, webseries_name)
    files = os.listdir(folder_path)
    for file in files:
        
            split_file = file.split('.')
            extension = split_file[-1]
            split_file=split_file[0].split(" ")
            split_file=split_file[2]
            list_of_all_numbers_in_filename = re.findall('[0-9]+',file)
            s_num = list_of_all_numbers_in_filename[0]
            s_num = str(int(s_num))
            e_num = list_of_all_numbers_in_filename[1]
            e_num = str(int(e_num))
            f_name = webseries_name + " - Season " + s_num.zfill(season_pad) + " Episode " + e_num.zfill(episode_pad) + "." + extension
            des=os.path.join(destiny,f_name)
            if not os.path.isdir(destiny): 
                os.chdir(dest)
                
                os.mkdir("Breaking Bad")
            if not os.path.isfile(des):
                
                os.chdir(destiny)
                
                with open(f_name, 'w') as file:
                  pass
        
    return
def rename_Gameofthrones(webseries_name,season_padding,episode_padding):
    folder1 = "correct_srt"
 
    if not os.path.exists(folder1): 
        os.makedirs(folder1)
    season_pad = (season_padding)
    episode_pad = (episode_padding)
    basepath = "wrong_srt"
    destiny=os.path.join(os.getcwd(),"correct_srt",webseries_name)

    path=os.path.join(os.getcwd(),"wrong_srt",webseries_name)
    dest=os.path.join(os.getcwd(),"correct_srt")
    folder_path = os.path.join(basepath, webseries_name)
    files = os.listdir(folder_path)
    for file in files:
        
            split_file = file.split('-')
            s_name = split_file[0]
            mid_split = split_file[1].split('x')
            mid_split = [ x.strip() for x in mid_split]
            s_num = mid_split[0]
            s_num = str(int(s_num))
            if len(split_file) > 3 and split_file[2].strip().isdigit():
                e_num = str(int(mid_split[1])).zfill(episode_pad) + "-" + str(int(split_file[2].strip())).zfill(episode_pad)
                end_split = split_file[3].split('.')
                e_name = end_split[0]
                e_name = e_name.strip()
                extension = file.split('.')[-1]
                f_name = s_name + "- " + "Season " + s_num.zfill(season_pad) + " Episode " + e_num + " - " + e_name + "." + extension
                des=os.path.join(destiny,f_name)
                if not os.path.isdir(destiny): 
                    os.chdir(dest)
                
                    os.mkdir("Game of Thrones")
                if not os.path.isfile(des):
                
                    os.chdir(destiny)
                
                    with open(f_name, 'w') as file:
                      pass
                

            else:
                e_num = mid_split[1]
                e_num = str(int(e_num))
                end_split = split_file[2].split('.')
                e_name = end_split[0]
                e_name = e_name.strip()
                extension = file.split('.')[-1]
    
                f_name = s_name + "- " + "Season " + s_num.zfill(season_pad) + " Episode " + e_num.zfill(episode_pad) + " - " + e_name + "." + extension
                des=os.path.join(destiny,f_name)
                if not os.path.isdir(destiny): 
                    os.chdir(dest)
                
                    os.mkdir("Game of Thrones")
                if not os.path.isfile(des):
                
                    os.chdir(destiny)
                
                    with open(f_name, 'w') as file:
                      pass
        
    return
def rename_Lucifer(webseries_name,season_padding,episode_padding):
    folder1 = "correct_srt"
 
    if not os.path.exists(folder1): 
        os.makedirs(folder1)
    season_pad = (season_padding)
    episode_pad = (episode_padding)
    basepath = "wrong_srt"
    destiny=os.path.join(os.getcwd(),"correct_srt",webseries_name)
    
    path=os.path.join(os.getcwd(),"wrong_srt",webseries_name)
    dest=os.path.join(os.getcwd(),"correct_srt")
    folder_path = os.path.join(basepath, webseries_name)
    files = os.listdir(folder_path)
    for file in files:
        
            split_file = file.split('-')
            s_name = split_file[0]
            mid_split = split_file[1].split('x')
            mid_split = [ x.strip() for x in mid_split]
            s_num = mid_split[0]
            s_num = str(int(s_num))
            if len(split_file) > 3 and split_file[2].strip().isdigit():
                e_num = str(int(mid_split[1])).zfill(episode_pad) + "-" + str(int(split_file[2].strip())).zfill(episode_pad)
                end_split = split_file[3].split('.')
                e_name = end_split[0]
                e_name = e_name.strip()
                extension = file.split('.')[-1]
                f_name = s_name + "- " + "Season " + s_num.zfill(season_pad) + " Episode " + e_num + " - " + e_name + "." + extension
                des=os.path.join(destiny,f_name)
                if not os.path.isdir(destiny): 
                    os.chdir(dest)
                
                    os.mkdir("Lucifer")
                if not os.path.isfile(des):
                
                    os.chdir(destiny)
                
                    with open(f_name, 'w') as file:
                      pass
                

            else:
                e_num = mid_split[1]
                e_num = str(int(e_num))
                end_split = split_file[2].split('.')
                e_name = end_split[0]
                e_name = e_name.strip()
                extension = file.split('.')[-1]
    
                f_name = s_name + "- " + "Season " + s_num.zfill(season_pad) + " Episode " + e_num.zfill(episode_pad) + " - " + e_name + "." + extension
                des=os.path.join(destiny,f_name)
                if not os.path.isdir(destiny): 
                    os.chdir(dest)
                
                    os.mkdir("Lucifer")
                if not os.path.isfile(des):
                
                    os.chdir(destiny)
                
                    with open(f_name, 'w') as file:
                      pass
        
    
def regex_renamer(): 

    # Taking input from the user
    print("1. Breaking Bad")
    print("2. Game of Thrones")
    print("3. Lucifer")

    webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
    season_padding = int(input("Enter the Season Number Padding: "))
    episode_padding = int(input("Enter the Episode Number Padding: "))

    if webseries_num==1 :  
        
            rename_Breaking_bad('Breaking Bad',season_padding,episode_padding)
         
    elif webseries_num==2: 
        
            rename_Gameofthrones('Game of Thrones',season_padding,episode_padding)
        
    elif webseries_num==3:  
        
            rename_Lucifer('Lucifer',season_padding,episode_padding)

    return 


regex_renamer()