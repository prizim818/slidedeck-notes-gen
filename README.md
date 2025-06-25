# slidedeck-notes-gen
Simple LLM script for reading in a slide deck (text-only) and using ChatGPT to generate summarized speaker's notes, then inserting those notes into the existing slide deck.

BASIC SETUP:
Modify the main python script to include your OpenAI API key. Then modify the two lines at the end of the script to specify your input and output folders. This script is designed to run batches of slide decks, hence the folder structure.

CUSTOMIZATION:
Modify the lines that contain the actual prompts to customize your output. I have included a generic prompt to start with, but I recommend changing it to include the topic, audience, level of complexity, and preferred tone to get best results. You can also modify the "chunk_size" variable to change how many slides it fits into one API call. Lower chunk sizes lead to more API calls but generally higher quality outputs. Default chunk size is 5.
