import pandas as pd
from gtts import gTTS



def TTS (input_file):

    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)
    
    for column in df.columns:
        # Iterate over each row in the column
        for index, cell_value in df[column].items():
            # Assuming the words are in the first column (index 0)
            print(f"Column: {column}, Row: {index}, Value: {cell_value}")

            # Specify the language of the word (e.g., "hi" for Hindi)
            language = column
            print (language)
            try:
                # Create a gTTS object and generate the speech
                tts = gTTS(text=str(cell_value), lang=language)
                # Save the MP3 file
                count = index + 1
                
                mp3_file = f"{language}_{count}.mp3"  # Customize the naming convention if needed
                tts.save(mp3_file)
                print(f"Generated MP3 for word '{language}_{cell_value}': {mp3_file}")      

                #break  # Translation successful, break out of the loop

            except Exception as e:
                    print(f"Audio Generation Failed for{language}{cell_value})")            

