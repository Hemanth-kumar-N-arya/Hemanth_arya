import tkinter as tk
import pandas as pd
from tkinter import filedialog
import openpyxl
import customtkinter
from openpyxl.styles import PatternFill
customtkinter.set_appearance_mode("Dark") 
customtkinter.set_default_color_theme("blue")  
root = customtkinter.CTk()


def browse_files():
    filename = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("Excel files", "*.xlsx*"), ("all files", "*.*")))
    if filename:
        create_checkbox_window(filename)


def create_checkbox_window(filename):
    checkbox_window = customtkinter.CTkToplevel(root)
    checkbox_window.title("DMS Data Annotation")
    Label = customtkinter.CTkLabel(checkbox_window, text=" DMS Annotation:", font=("Arial", 16))
    Label.grid(row=0, column=2, padx=0, pady=10, sticky="w")


    option1_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeDetected_left_after", variable=option1_var).grid(row=1, column=1, padx=5, pady=10, sticky="w")
    option2_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeDetectedConfidence_left_after", variable=option2_var).grid(row=2, column=1, padx=5, pady=10, sticky="w")
    option3_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeOpenLevel_left_after", variable=option3_var).grid(row=3, column=1, padx=5, pady=10, sticky="w")
    option4_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeOpenLevelConfidence_left_after", variable=option4_var).grid(row=4, column=1, padx=5, pady=10,  sticky="w")
    option5_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeDetected_right_after", variable=option5_var).grid(row=5, column=1, padx=5, pady=10,  sticky="w")
    option6_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeDetectedConfidence_right_after", variable=option6_var).grid(row=6, column=1, padx=5, pady=10,  sticky="w")
    option7_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeOpenLevel_right_after", variable=option7_var).grid(row=7, column=1, padx=5, pady=10,  sticky="w")
    option8_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeOpenLevelConfidence_right_after", variable=option8_var).grid(row=8, column=1, padx=5, pady=10,  sticky="w")
    option9_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="head_found_after", variable=option9_var).grid(row=9, column=1, padx=5, pady=10,  sticky="w")
    option10_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="head_found_conf_after", variable=option10_var).grid(row=10, column=1, padx=5, pady=10,  sticky="w")
    option11_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="face_visibility_after", variable=option11_var).grid(row=11, column=1, padx=5, pady=10,  sticky="w")
    option12_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="face_visibility_conf_after", variable=option12_var).grid(row=1, column=3, padx=5, pady=10, sticky="w")
    option13_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="glass_status_after", variable=option13_var).grid(row=2, column=3, padx=5, pady=10, sticky="w")
    option14_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="IR_glass_status_after", variable=option14_var).grid(row=3, column=3, padx=5, pady=10, sticky="w")
    option15_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="empty_driver_seat_after", variable=option15_var).grid(row=4, column=3, padx=5, pady=10,  sticky="w")
    option16_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="fake_detection_after", variable=option16_var).grid(row=5, column=3, padx=5, pady=10,  sticky="w")
    option17_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="camera_blindness_after", variable=option17_var).grid(row=6, column=3, padx=5, pady=10,  sticky="w")
    option18_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyes_on_road_after", variable=option18_var).grid(row=7, column=3, padx=5, pady=10,  sticky="w")
    option19_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="driver_change_after", variable=option19_var).grid(row=8, column=3, padx=5, pady=10,  sticky="w")
    option20_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeOccluded_left_after", variable=option20_var).grid(row=9, column=3, padx=5, pady=10,  sticky="w")
    option21_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="eyeOccluded_right_after", variable=option21_var).grid(row=10, column=3, padx=5, pady=10,  sticky="w")
    option22_var = tk.IntVar()
    customtkinter.CTkCheckBox(checkbox_window, text="without_glasses", variable=option22_var).grid(row=11, column=3, padx=5, pady=10,  sticky="w")

    submit_button = customtkinter.CTkButton(checkbox_window, text="Submit", command=lambda: write_to_file(option1_var.get(), option2_var.get(), option3_var.get(),option4_var.get(),option5_var.get(), option6_var.get(), option7_var.get(),option8_var.get(),option9_var.get(), option10_var.get(),option11_var.get(), option12_var.get(), option13_var.get(),option14_var.get(),option15_var.get(), option16_var.get(), option17_var.get(),option18_var.get(),option19_var.get(), option20_var.get(),option21_var.get(),option22_var.get(), filename))
    submit_button.grid(row=14, column=2, padx=5, pady=10, sticky="w")

def left_eye(option1_value, option2_value, option3_value,option4_value, filename):
    #sheet_name= "Sheet1"
    #column_name = ["Frame_No","eyeDetected_left_after","eyeDetectedConfidence_left_after","eyeOpenLevel_left_after","eyeOpenLevelConfidence_left_after","eyeDetected_right_after","eyeDetectedConfidence_right_after","eyeOpenLevel_right_after","eyeOpenLevelConfidence_right_after","head_found_after","head_found_conf_after","face_visibility_after","face_visibility_conf_after","glass_status_after","IR_glass_status_after","empty_driver_seat_after","fake_detection_after","camera_blindness_after","eyes_on_road_after","driver_change_after","eyeOccluded_left_after","eyeOccluded_right_after"]
    #df = pd.read_excel(filename, sheet_name=sheet_name)

    if (option1_value == 1 and option2_value == 1 and option3_value == 1 and option4_value == 1):
        #df = pd.DataFrame(df)
        df['left_verify'] = 'pass'
        df.loc[
            (df['eyeDetected_left_after'] == 'eyeDetected') &
                ((df['eyeDetectedConfidence_left_after'] == 'notLabeled') |
                (df['eyeOpenLevel_left_after'] == 'notLabeled') |
                (df['eyeOpenLevelConfidence_left_after'] == 'notLabeled')
            ), 'left_verify'] = 'check'
        df.loc[
            (df['eyeDetected_left_after'] == 'eyeNotDetected') &
                (df['eyeDetectedConfidence_left_after'] == 'sure') &
                ((df['eyeOpenLevel_left_after'] != 'notLabeled') |
                (df['eyeOpenLevelConfidence_left_after'] != 'notLabeled')
            ), 'left_verify'] = 'check'
        df.loc[
            (((df['eyeDetected_left_after'] == 'eyeNotDetected') |
                (df['eyeDetected_left_after'] == 'eyeDetected')) &
                (df['eyeDetectedConfidence_left_after'] == 'unsure') &
                ((df['eyeOpenLevel_left_after'] == 'notLabeled') |
                (df['eyeOpenLevelConfidence_left_after'] == 'notLabeled'))
            ), 'left_verify'] = 'check'
        df.loc[
            (((df['eyeOpenLevel_left_after'] == 'eyeNotClosed') |
                 (df['eyeOpenLevel_left_after'] == 'eyeClosed')) &
                  ((df['eyeDetected_left_after'] == 'notLabeled') |
                   (df['eyeDetectedConfidence_left_after'] == 'notLabeled') |
                   (df['eyeOpenLevelConfidence_left_after'] == 'notLabeled'))
            ), 'left_verify'] = 'check'
        df.loc[
            (((df['eyeDetectedConfidence_left_after'] == 'sure') |
                   (df['eyeDetectedConfidence_left_after'] == 'unsure')) &
                   (df['eyeDetected_left_after'] == 'notLabeled')
            ), 'left_verify'] = 'check'
        df.loc[
            (((df['eyeOpenLevelConfidence_left_after'] == 'sure') |
                    (df['eyeOpenLevelConfidence_left_after'] == 'unsure')) &
                   ((df['eyeDetectedConfidence_left_after'] == 'notLabeled') |
                    (df['eyeOpenLevel_left_after'] == 'notLabeled') |
                    (df['eyeDetected_left_after'] == 'notLabeled'))
             ), 'left_verify'] = 'check'
        df.loc[
            ((df['eyeDetected_left_after'] == 'eyeNotDetected') &
                  (df['eyeDetectedConfidence_left_after'] == 'notLabeled')
            ), 'left_verify'] = 'check'
        df.loc[
            ((df['eyeDetected_left_after'] == 'notLabeled') &
                 (df['eyeDetectedConfidence_left_after'] == 'notLabeled') &
                 (df['eyeOpenLevel_left_after'] == 'notLabeled') &
                 (df['eyeOpenLevelConfidence_left_after'] == 'notLabeled')
            ), 'left_verify'] = 'check'
        print(df['left_verify'])
    
    elif (option1_value == 1 and option2_value == 1 and option3_value == 0  and  option4_value == 0):
        df['left_verify'] = 'pass'
        df.loc[((df['eyeDetected_left_after'] != 'eyeDetected') | (df['eyeDetectedConfidence_left_after'] == 'notLabeled')) | 
               ((df['eyeOpenLevel_left_after'] != 'notLabeled') | (df['eyeOpenLevelConfidence_left_after'] != 'notLabeled')) , 'left_verify'] ='check'
        print(df['left_verify'])
        #df[['Frame_No','left_verify']].to_excel(filename, sheet_name = 'Error', index =False)
        
    elif (option1_value == 0 and option2_value == 0 and option3_value == 1  and  option4_value == 1):
        df['left_verify'] = 'pass'
        df.loc[((df['eyeDetected_left_after'] != 'notLabeled') & (df['eyeDetectedConfidence_left_after'] != 'notLabeled')) | 
               ((df['eyeOpenLevel_left_after'] == 'notLabeled') & (df['eyeOpenLevelConfidence_left_after'] == 'notLabeled')) , 'left_verify'] ='check'
        print(df['left_verify'])
        
    elif (option1_value == 1 and option2_value == 0 and option3_value == 0  and  option4_value == 0):
        df['left_verify'] = 'pass'
        df.loc[(df['eyeDetected_left_after'] != 'eyeDetected') | ((df['eyeDetectedConfidence_left_after'] != 'notLabeled') | 
               (df['eyeOpenLevel_left_after'] != 'notLabeled') | (df['eyeOpenLevelConfidence_left_after'] != 'notLabeled')) , 'left_verify'] ='check'
        print(df['left_verify'])
        
    elif (option1_value == 1 and option2_value == 0 and option3_value == 1  and  option4_value == 0):
        df['left_verify'] = 'pass'
        df.loc[((df['eyeDetected_left_after'] != 'eyeDetected') & (df['eyeOpenLevel_left_after'] == 'notLabeled')) | 
               ((df['eyeDetectedConfidence_left_after'] != 'notLabeled') & (df['eyeOpenLevelConfidence_left_after'] != 'notLabeled')) , 'left_verify'] ='check'
        print(df['left_verify'])
        
    elif (option1_value == 0 and option2_value == 0 and option3_value == 1  and  option4_value == 0):
        df['left_verify'] = 'pass'
        df.loc[((df['eyeDetected_left_after'] != 'notLabeled') | (df['eyeDetectedConfidence_left_after'] == 'notLabeled') | (df['eyeOpenLevelConfidence_left_after'] != 'notLabeled')) | (df['eyeOpenLevel_left_after'] == 'notLabeled') , 'left_verify'] ='check'
        print(df['left_verify'])
        
    elif (option1_value == 0 and option2_value == 0 and option3_value == 0  and  option4_value == 0):
        df['left_verify'] = 'pass'
        df.loc[((df['eyeDetected_left_after'] != 'notLabeled') | (df['eyeDetectedConfidence_left_after'] != 'notLabeled')) | 
               ((df['eyeOpenLevel_left_after'] != 'notLabeled') | (df['eyeOpenLevelConfidence_left_after'] != 'notLabeled')) , 'left_verify'] ='check'
        print(df['left_verify'])
        
def right_eye(option5_value, option6_value, option7_value,option8_value, filename):
    #sheet_name= "Sheet1"
    #column_name = ["Frame_No","eyeDetected_left_after","eyeDetectedConfidence_left_after","eyeOpenLevel_left_after","eyeOpenLevelConfidence_left_after","eyeDetected_right_after","eyeDetectedConfidence_right_after","eyeOpenLevel_right_after","eyeOpenLevelConfidence_right_after","head_found_after","head_found_conf_after","face_visibility_after","face_visibility_conf_after","glass_status_after","IR_glass_status_after","empty_driver_seat_after","fake_detection_after","camera_blindness_after","eyes_on_road_after","driver_change_after","eyeOccluded_left_after","eyeOccluded_right_after"]
    #df = pd.read_excel(filename, sheet_name=sheet_name)

    if (option5_value == 1 and option6_value == 1 and option7_value == 1 and option8_value == 1):
        #df = pd.DataFrame(df)
        df['right_verify'] = 'pass'
        df.loc[
            (df['eyeDetected_right_after'] == 'eyeDetected') &
                ((df['eyeDetectedConfidence_right_after'] == 'notLabeled') |
                (df['eyeOpenLevel_right_after'] == 'notLabeled') |
                (df['eyeOpenLevelConfidence_right_after'] == 'notLabeled')
            ), 'right_verify'] = 'check'
        df.loc[
            (df['eyeDetected_right_after'] == 'eyeNotDetected') &
                (df['eyeDetectedConfidence_right_after'] == 'sure') &
                ((df['eyeOpenLevel_right_after'] != 'notLabeled') |
                (df['eyeOpenLevelConfidence_right_after'] != 'notLabeled')
            ), 'right_verify'] = 'check'
        df.loc[
            (((df['eyeDetected_right_after'] == 'eyeNotDetected') |
                (df['eyeDetected_right_after'] == 'eyeDetected')) &
                (df['eyeDetectedConfidence_right_after'] == 'unsure') &
                ((df['eyeOpenLevel_right_after'] == 'notLabeled') |
                (df['eyeOpenLevelConfidence_right_after'] == 'notLabeled'))
            ), 'right_verify'] = 'check'
        df.loc[
            (((df['eyeOpenLevel_right_after'] == 'eyeNotClosed') |
                 (df['eyeOpenLevel_right_after'] == 'eyeClosed')) &
                  ((df['eyeDetected_right_after'] == 'notLabeled') |
                   (df['eyeDetectedConfidence_right_after'] == 'notLabeled') |
                   (df['eyeOpenLevelConfidence_right_after'] == 'notLabeled'))
            ), 'right_verify'] = 'check'
        df.loc[
            (((df['eyeDetectedConfidence_right_after'] == 'sure') |
                   (df['eyeDetectedConfidence_right_after'] == 'unsure')) &
                   (df['eyeDetected_right_after'] == 'notLabeled')
            ), 'right_verify'] = 'check'
        df.loc[
            (((df['eyeOpenLevelConfidence_right_after'] == 'sure') |
                    (df['eyeOpenLevelConfidence_right_after'] == 'unsure')) &
                   ((df['eyeDetectedConfidence_right_after'] == 'notLabeled') |
                    (df['eyeOpenLevel_right_after'] == 'notLabeled') |
                    (df['eyeDetected_right_after'] == 'notLabeled'))
             ), 'right_verify'] = 'check'
        df.loc[
            ((df['eyeDetected_right_after'] == 'eyeNotDetected') &
                  (df['eyeDetectedConfidence_right_after'] == 'notLabeled')
            ), 'right_verify'] = 'check'
        df.loc[
            ((df['eyeDetected_right_after'] == 'notLabeled') &
                 (df['eyeDetectedConfidence_right_after'] == 'notLabeled') &
                 (df['eyeOpenLevel_right_after'] == 'notLabeled') &
                 (df['eyeOpenLevelConfidence_right_after'] == 'notLabeled')
            ), 'right_verify'] = 'check'
        print(df['right_verify'])
        
    elif (option5_value == 1 and option6_value == 1 and option7_value == 0 and option8_value == 0):
        df['right_verify'] = 'pass'
        df.loc[((df['eyeDetected_right_after'] != 'eyeDetected') | (df['eyeDetectedConfidence_right_after'] == 'notLabeled')) | 
               ((df['eyeOpenLevel_right_after'] != 'notLabeled') | (df['eyeOpenLevelConfidence_right_after'] != 'notLabeled')) , 'right_verify'] ='check'
        print(df['right_verify'])
        
    elif (option5_value == 0 and option6_value == 0 and option7_value == 1 and option8_value == 1):
        df['right_verify'] = 'pass'
        df.loc[((df['eyeDetected_right_after'] != 'notLabeled') & (df['eyeDetectedConfidence_right_after'] != 'notLabeled')) | 
               ((df['eyeOpenLevel_right_after'] == 'notLabeled') & (df['eyeOpenLevelConfidence_right_after'] == 'notLabeled')) , 'right_verify'] ='check'
        print(df['right_verify'])
        
    elif (option5_value == 1 and option6_value == 0 and option7_value == 0 and option8_value == 0):
        df['right_verify'] = 'pass'
        df.loc[(df['eyeDetected_right_after'] != 'eyeDetected') | ((df['eyeDetectedConfidence_right_after'] != 'notLabeled') | 
               (df['eyeOpenLevel_right_after'] != 'notLabeled') | (df['eyeOpenLevelConfidence_right_after'] != 'notLabeled')) , 'right_verify'] ='check'
        print(df['right_verify'])
        
    elif (option5_value == 1 and option6_value == 0 and option7_value == 1 and option8_value == 0):
        df['right_verify'] = 'pass'
        df.loc[((df['eyeDetected_right_after'] != 'eyeDetected') & (df['eyeOpenLevel_right_after'] == 'notLabeled')) | 
               ((df['eyeDetectedConfidence_right_after'] != 'notLabeled') & (df['eyeOpenLevelConfidence_right_after'] != 'notLabeled')) , 'right_verify'] ='check'
        print(df['right_verify'])
        
    elif (option5_value == 0 and option6_value == 0 and option7_value == 1 and option8_value == 0):
        df['right_verify'] = 'pass'
        df.loc[((df['eyeDetected_right_after'] != 'notLabeled') | (df['eyeDetectedConfidence_right_after'] == 'notLabeled') | (df['eyeOpenLevelConfidence_right_after'] != 'notLabeled')) | (df['eyeOpenLevel_right_after'] == 'notLabeled') , 'right_verify'] ='check'
        print(df['right_verify'])
    
    elif (option5_value == 0 and option6_value == 0 and option7_value == 0 and option8_value == 0):
        df['right_verify'] = 'pass'
        df.loc[((df['eyeDetected_right_after'] != 'notLabeled') | (df['eyeDetectedConfidence_right_after'] != 'notLabeled')) | 
               ((df['eyeOpenLevel_right_after'] != 'notLabeled') | (df['eyeOpenLevelConfidence_right_after'] != 'notLabeled')) , 'right_verify'] ='check'
        print(df['right_verify'])
    
def head_found(option9_value, option10_value, filename):
    #sheet_name= "Sheet1"
    #column_name = ["Frame_No","eyeDetected_left_after","eyeDetectedConfidence_left_after","eyeOpenLevel_left_after","eyeOpenLevelConfidence_left_after","eyeDetected_right_after","eyeDetectedConfidence_right_after","eyeOpenLevel_right_after","eyeOpenLevelConfidence_right_after","head_found_after","head_found_conf_after","face_visibility_after","face_visibility_conf_after","glass_status_after","IR_glass_status_after","empty_driver_seat_after","fake_detection_after","camera_blindness_after","eyes_on_road_after","driver_change_after","eyeOccluded_left_after","eyeOccluded_right_after"]
    #df = pd.read_excel(filename, sheet_name=sheet_name)

    if (option9_value == 1 and option10_value == 1) :
        df[['Head_Found']] = 'Pass'
        df.loc[(df['head_found_after'] != 'headFound')  |(df['head_found_conf_after'] == 'notLabeled'),'Head_Found'] = 'check'
        df.loc[
            (df['head_found_after'] == 'notLabeled') | (df['head_found_conf_after'] == 'notLabeled'),
            'Head_Found'
        ] = 'check'
        print(df[['Head_Found']])

        #df[['Frame_No','Head_Found']].to_excel(filename, sheet_name = 'Error', index =False)
        
    elif (option9_value == 1 and option10_value == 0) :
        df['Head_Found'] = 'pass'
        df.loc[(df['head_found_after'] == 'notLabeled') | (df['head_found_conf_after'] != 'notLabeled') , 'Head_Found'] ='check'
        print(df['Head_Found'])
        

        
    elif (option9_value == 0 and option10_value == 0) :
        df['Head_Found'] = 'pass'
        df.loc[(df['head_found_after'] != 'notLabeled') | (df['head_found_conf_after'] != 'notLabeled') , 'Head_Found'] ='check'
        print(df['Head_Found'])
    

def face_visibility(option11_value, option12_value, filename):
    #sheet_name= "Sheet1"
    #column_name = ["Frame_No","eyeDetected_left_after","eyeDetectedConfidence_left_after","eyeOpenLevel_left_after","eyeOpenLevelConfidence_left_after","eyeDetected_right_after","eyeDetectedConfidence_right_after","eyeOpenLevel_right_after","eyeOpenLevelConfidence_right_after","head_found_after","head_found_conf_after","face_visibility_after","face_visibility_conf_after","glass_status_after","IR_glass_status_after","empty_driver_seat_after","fake_detection_after","camera_blindness_after","eyes_on_road_after","driver_change_after","eyeOccluded_left_after","eyeOccluded_right_after"]
    #df = pd.read_excel(filename, sheet_name=sheet_name)

    if (option11_value == 1 and option12_value == 1) :
        df['FACE'] = 'pass'
        df.loc[(df['face_visibility_after'] != 'faceVisible')  |(df['face_visibility_conf_after'] == 'notLabeled'),'FACE'] = 'check'
        df.loc[
            (df['face_visibility_after'] == 'notLabeled') | (df['face_visibility_conf_after'] == 'notLabeled'),
             'FACE'
        ] = 'check'
        print(df['FACE'])
        
    elif (option11_value == 1 and option12_value == 0) :
        df['FACE'] = 'pass'
        df.loc[(df['face_visibility_after'] == 'notLabeled') | (df['face_visibility_conf_after'] != 'notLabeled') , 'FACE'] ='check'
        print(df['FACE'])
        
    elif (option11_value == 0 and option12_value == 1) :
        df['FACE'] = 'pass'
        df.loc[(df['face_visibility_after'] != 'notLabeled') | (df['face_visibility_conf_after'] == 'notLabeled') , 'FACE'] ='check'
        print(df['FACE'])
        
    elif (option11_value == 0 and option12_value == 0) :
        df['FACE'] = 'pass'
        df.loc[(df['face_visibility_after'] != 'notLabeled') | (df['face_visibility_conf_after'] != 'notLabeled') , 'FACE'] ='check'
        print(df['FACE'])
    
def glasses(option13_value, option14_value, option22_value, filename):
    if option13_value == 1 and option14_value == 1 and option22_value == 0:
        df['GLASS'] = 'pass'
        df['IR_GLASS'] = 'pass'
        df.loc[(df['glass_status_after'] != 'glassesDetected'), 'GLASS'] = 'check'
        df.loc[(df['IR_glass_status_after'] != 'irGlassesDetected'), 'IR_GLASS'] = 'check'
        print(df[['GLASS', 'IR_GLASS']])
        print("selected")
    elif option13_value == 1 and option14_value == 0 and option22_value == 0:
        df['GLASS'] = 'pass'
        df['IR_GLASS'] = 'pass'
        df.loc[(df['glass_status_after'] != 'glassesDetected'), 'GLASS'] = 'check'
        df.loc[(df['IR_glass_status_after'] != 'irGlassesNotDetected'), 'IR_GLASS'] = 'check'
        print(df[['GLASS', 'IR_GLASS']])
        print("selected")
    elif option13_value == 0 and option14_value == 1 and option22_value == 0:
        df['GLASS'] = 'pass'
        df['IR_GLASS'] = 'pass'
        df.loc[(df['glass_status_after'] != 'glassesDetected'), 'GLASS'] = 'check'
        df.loc[(df['IR_glass_status_after'] != 'irGlassesDetected'), 'IR_GLASS'] = 'check'
        print(df[['GLASS', 'IR_GLASS']])
        print("selected")
    elif option13_value == 0 and option14_value == 0 and option22_value == 1:
        df['GLASS'] = 'pass'
        df['IR_GLASS'] = 'pass'
        df.loc[(df['glass_status_after'] != 'glassesNotDetected'), 'GLASS'] = 'check'
        df.loc[(df['IR_glass_status_after'] != 'irGlassesNotDetected'), 'IR_GLASS'] = 'check'
        print(df[['GLASS', 'IR_GLASS']])
    else:
        df['GLASS'] = 'pass'
        df['IR_GLASS'] = 'pass'
        df.loc[(df['glass_status_after'] != 'notLabeled'), 'GLASS'] = 'check'
        df.loc[(df['IR_glass_status_after'] != 'notLabeled'), 'IR_GLASS'] = 'check'
        print(df[['GLASS', 'IR_GLASS']])
    
    #df.to_excel(filename, columns = ['Frame_No', 'GLASS', 'IR_GLASS'], sheet_name='Error', index=False)

def empty_driver_seat(option15_value, filename):
    
    if option15_value == 1 :
        
        df['SEAT_OCCUPIED'] ='pass'
        
        df.loc[(df['empty_driver_seat_after'] != 'notLabeled') &
               (df['camera_blindness_after'] == 'cameraCovered') , 'SEAT_OCCUPIED'] = 'check'
        df.loc[(df['empty_driver_seat_after'] == 'notLabeled') &
               (df['camera_blindness_after'] == 'cameraNotCovered') , 'SEAT_OCCUPIED'] = 'check'

        print(df['SEAT_OCCUPIED'])
        
    elif option15_value == 0 :
        df['SEAT_OCCUPIED'] ='pass'
        df.loc[(df['empty_driver_seat_after'] != 'notLabeled') , 'SEAT_OCCUPIED'] ='check'
        print(df['SEAT_OCCUPIED'])

        
def fake_detection(option16_value, filename):

    if option16_value == 1 :
        df['FAKE_DETECTION'] ='pass'
        df.loc[(df['fake_detection_after'] != 'notLabeled') &
            (df['camera_blindness_after'] == 'cameraCovered') , 'FAKE_DETECTION'] = 'check'
        df.loc[(df['fake_detection_after'] == 'notLabeled') &
            (df['camera_blindness_after'] == 'cameraNotCovered') , 'FAKE_DETECTION'] = 'check'
        print(df['FAKE_DETECTION'])
        
    elif option16_value == 0 :
        df['FAKE_DETECTION'] ='pass'
        df.loc[(df['fake_detection_after'] != 'notLabeled') , 'FAKE_DETECTION'] = 'check'
        print(df['FAKE_DETECTION'])


def camera_covered(option17_value, filename):

    if option17_value == 1 :
        df['CAMERA_COVERED'] ='pass'
        df.loc[(df['camera_blindness_after'] == 'notLabeled') , 'CAMERA_COVERED'] = 'check'
        df.loc[(df['camera_blindness_after'] == 'cameraCovered')&
               ((df['head_found_after'] == 'headFound')|(df['face_visibility_after'] == 'faceVisible')|
                (df['glass_status_after'] == 'glassesDetected')|
                (df['IR_glass_status_after'] == 'irGlassesDetected')|
                (df['empty_driver_seat_after'] == 'seatOccupied')),'CAMERA_COVERED'] = 'check'
        
    elif option17_value == 0 :
        df['CAMERA_COVERED'] ='pass'
        df.loc[(df['camera_blindness_after'] != 'notLabeled') , 'CAMERA_COVERED'] = 'check'
        print(df['CAMERA_COVERED'])


def eyes_on_road(option18_value, filename):

    if option18_value == 1 :
        df['EYES_ON_ROAD'] ='pass'
        df.loc[(df['eyes_on_road_after'] == 'notLabeled') , 'EYES_ON_ROAD'] = 'check'
        print(df['EYES_ON_ROAD'])
        
    elif option18_value == 0 :
        df['EYES_ON_ROAD'] ='pass'
        df.loc[(df['eyes_on_road_after'] != 'notLabeled') , 'EYES_ON_ROAD'] = 'check'
        print(df['EYES_ON_ROAD'])
    

def driver_change(option19_value, filename):

    if option19_value == 1 :
        df['DRIVER_CHANGE'] ='pass'
        df.loc[(df['driver_change_after'] == 'notLabeled') , 'DRIVER_CHANGE'] = 'check'
        print(df['DRIVER_CHANGE'])
        
    elif option19_value == 0 :
        df['DRIVER_CHANGE'] ='pass'
        df.loc[(df['driver_change_after'] != 'notLabeled') , 'DRIVER_CHANGE'] = 'check'
        print(df['DRIVER_CHANGE'])


def eye_oclusion(option20_value,option21_value, filename):

    if option20_value == 1 and option21_value == 1:
        df['EYE_OCCLUSION'] ='pass'
        df.loc[
            (df['eyeOccluded_left_after'] == 'notLabeled') | (df['eyeOccluded_right_after'] == 'notLabeled'),
                'EYE_OCCLUSION'] = 'check'
        df.loc[
            (df['eyeDetected_left_after']== 'eyeDetected') & (df['eyeDetectedConfidence_left_after'] == 'sure')&
            (df['eyeOccluded_left_after'] == 'eyeOccluded'),'EYE_OCCLUSION'] = 'check'
        df.loc[
            (df['eyeDetected_right_after']== 'eyeDetected') & (df['eyeDetectedConfidence_right_after'] == 'sure')&
            (df['eyeOccluded_right_after'] == 'eyeOccluded'),'EYE_OCCLUSION'] = 'check'
        df.loc[
            (df['eyeDetected_left_after']== 'eyeNotDetected') & (df['eyeDetectedConfidence_left_after'] == 'sure')&
            (df['eyeOccluded_left_after'] != 'fullyVisible'),'EYE_OCCLUSION'] = 'check'
        df.loc[
            (df['eyeDetected_right_after']== 'eyeNotDetected') & (df['eyeDetectedConfidence_right_after'] == 'sure')&
            (df['eyeOccluded_right_after'] != 'fullyVisible'),'EYE_OCCLUSION'] = 'check'
        print(df['EYE_OCCLUSION'])

        
    elif option20_value == 0 and option21_value == 0:
        df['EYE_OCCLUSION'] ='pass'
        df.loc[
            (df['eyeOccluded_left_after'] != 'notLabeled') & (df['eyeOccluded_right_after'] != 'notLabeled'),
            'EYE_OCCLUSION'
        ] = 'check'
        print(df['EYE_OCCLUSION'])
        #df[['Frame_No','left_verify','right_verify','Head_Found','FACE','GLASS' ,'IR_GLASS','SEAT_OCCUPIED','FAKE_DETECTION', 'CAMERA_COVERED','EYES_ON_ROAD','DRIVER_CHANGE','EYE_OCCLUSION' ]].to_excel('C:/Users/bj5pjj/Downloads/AP399_003_GO7_TS046_20230314_MAX_static_X084_C0_0000_final - Copy (3).xlsx')
    #else:
        #df.to_excel(filename, columns =['Frame_No','left_verify','right_verify','Head_Found','FACE','GLASS' ,'IR_GLASS','SEAT_OCCUPIED','FAKE_DETECTION', 'CAMERA_COVERED','EYES_ON_ROAD','DRIVER_CHANGE','EYE_OCCLUSION'] , sheet_name = 'Sheet1' ,index =False)

        
def write_to_file(option1_value, option2_value, option3_value, option4_value, option5_value, option6_value, option7_value, option8_value, option9_value, option10_value, option11_value, option12_value, option13_value, option14_value, option15_value, option16_value, option17_value, option18_value, option19_value, option20_value, option21_value, option22_value, filename):

    global df
    sheet_name = "Sheet1"
    column_name = ["Frame_No", "eyeDetected_left_after", "eyeDetectedConfidence_left_after",
                   "eyeOpenLevel_left_after", "eyeOpenLevelConfidence_left_after", "eyeDetected_right_after",
                   "eyeDetectedConfidence_right_after", "eyeOpenLevel_right_after",
                   "eyeOpenLevelConfidence_right_after", "head_found_after", "head_found_conf_after",
                   "face_visibility_after", "face_visibility_conf_after", "glass_status_after",
                   "IR_glass_status_after", "empty_driver_seat_after", "fake_detection_after",
                   "camera_blindness_after", "eyes_on_road_after", "driver_change_after",
                   "eyeOccluded_left_after", "eyeOccluded_right_after"]
    df1 = pd.read_excel(filename, sheet_name=sheet_name)
    df = pd.DataFrame(df1)

    if option1_value == 1:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option2_value == 1:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option3_value == 1:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option4_value == 1:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option5_value == 1:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option6_value == 1:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option7_value == 1:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option8_value == 1:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option9_value == 1:
        head_found(option9_value, option10_value, filename)

    if option10_value == 1:
        head_found(option9_value, option10_value, filename)

    if option11_value == 1:
        face_visibility(option11_value, option12_value, filename)

    if option12_value == 1:
        face_visibility(option11_value, option12_value, filename)

    if option13_value == 1:
        glasses(option13_value, option14_value, option22_value, filename)

    if option14_value == 1:
        glasses(option13_value, option14_value, option22_value, filename)

    if option22_value == 1:
        glasses(option13_value, option14_value, option22_value, filename)

    if option15_value == 1:
        empty_driver_seat(option15_value, filename)

    if option16_value == 1:
        fake_detection(option16_value, filename)

    if option17_value == 1:
        camera_covered(option17_value, filename)

    if option18_value == 1:
        eyes_on_road(option18_value, filename)

    if option19_value == 1:
        driver_change(option19_value, filename)

    if option20_value == 1:
        eye_oclusion(option20_value,option21_value, filename)

    if option21_value == 1:
        eye_oclusion(option20_value,option21_value, filename)

    if option1_value == 0:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option2_value == 0:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option3_value == 0:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option4_value == 0:
        left_eye(option1_value, option2_value, option3_value, option4_value, filename)

    if option5_value == 0:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option6_value == 0:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option7_value == 0:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option8_value == 0:
        right_eye(option5_value, option6_value, option7_value, option8_value, filename)

    if option9_value == 0:
        head_found(option9_value, option10_value, filename)

    if option10_value == 0:
        head_found(option9_value, option10_value, filename)

    if option11_value == 0:
        face_visibility(option11_value, option12_value, filename)

    if option12_value == 0:
        face_visibility(option11_value, option12_value, filename)

    if option13_value == 0:
        glasses(option13_value, option14_value, option22_value, filename)

    if option14_value == 0:
        glasses(option13_value, option14_value, option22_value, filename)

    if option22_value == 0:
        glasses(option13_value, option14_value, option22_value, filename)

    if option15_value == 0:
        empty_driver_seat(option15_value, filename)

    if option16_value == 0:
        fake_detection(option16_value, filename)

    if option17_value == 0:
        camera_covered(option17_value, filename)

    if option18_value == 0:
        eyes_on_road(option18_value, filename)

    if option19_value == 0:
        driver_change(option19_value, filename)

    if option20_value == 0:
        eye_oclusion(option20_value,option21_value, filename)

    if option21_value == 0:
        eye_oclusion(option20_value,option21_value, filename)

    
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    
    #ptint of excel sheet 1
    df1.to_excel(writer, sheet_name='Sheet1', index=False)
    
    


   #print of sheet ERROR of Pass n check
    df[['Frame_No', 'left_verify', 'right_verify', 'Head_Found', 'FACE', 'GLASS', 'IR_GLASS', 'SEAT_OCCUPIED',
        'FAKE_DETECTION', 'CAMERA_COVERED', 'EYES_ON_ROAD', 'DRIVER_CHANGE', 'EYE_OCCLUSION']].to_excel(writer,
                                                                                                        sheet_name='error',
                                                                                                         index=False)


    
    df2 = [
        len(df["Frame_No"]),
        len(df[df["left_verify"] == "check"]),
        len(df[df["right_verify"] == "check"]),
        len(df[df["Head_Found"] == "check"]),
        len(df[df["FACE"] == "check"]),
        len(df[df["GLASS"] == "check"]),
        len(df[df["IR_GLASS"] == "check"]),
        len(df[df["SEAT_OCCUPIED"] == "check"]),
        len(df[df["FAKE_DETECTION"] == "check"]),
        len(df[df["CAMERA_COVERED"] == "check"]),
        len(df[df["EYES_ON_ROAD"] == "check"]),
        len(df[df["DRIVER_CHANGE"] == "check"]),
        len(df[df["EYE_OCCLUSION"] == "check"])
    ]

    df2 = pd.DataFrame([df2],
                       columns=['Frame_No', 'left_verify', 'right_verify', 'Head_Found', 'FACE', 'GLASS', 'IR_GLASS',
                                'SEAT_OCCUPIED', 'FAKE_DETECTION', 'CAMERA_COVERED', 'EYES_ON_ROAD', 'DRIVER_CHANGE',
                                'EYE_OCCLUSION'])

    
    
    df2.to_excel(writer, sheet_name='error',startrow=0, startcol=14, index=False)

    sheet = writer.book['error']
    light_pink_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    columns_to_highlight = ['B', 'C', 'D','E', 'F', 'G', 'H','I', 'J', 'K', 'L', 'M' ] 
    for column in columns_to_highlight:
        for cell in sheet[column]:
            if cell.value == 'check':
                cell.fill = light_pink_fill

    writer.close()
    
    print(df2)



root.title("File Explorer")

browse_button = customtkinter.CTkButton(root, text="Browse Files", command=browse_files)
browse_button.grid(row=12, column=1, padx=10, pady=10, sticky="w")

root.mainloop()
