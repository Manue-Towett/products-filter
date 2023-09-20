# products-filter
Filters products based on the given settings

# Requires
- python 3.11+

# Settings
- There are two configuration files:
    - settings:
        - use this to set:
            - input - the folder where the input excel files are located
            - output - the folder where the processed files will go to
            - AmazonPrice - the minimum amazon price
            - ROI - the minimum ROI
            - Rating - the minimum rating
            - ReviewCount - the minimum reviews
            - OfferCount - the minimum offers
            - offers_0_availability - availability
    
    - blacklisted:
        - write blacklisted product in this file, each on its own line with no empty lines

# Usage
- Open the command prompt / terminal
- cd into the project folder / directory
- If running for the first time, first install dependencies using the command: 
    ```pip install -r requirements.txt```
- To run the script, use the commands:
    - For Linux/MacOS: 
        ```python3 main.py```
    - For windows: 
        ```python main.py```