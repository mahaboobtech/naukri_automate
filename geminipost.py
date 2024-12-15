import sec
import google.generativeai as genai



def load_user_data(file_path):
    """
    Load user data from a text file to provide personalized answers.
    Each line in the file should have the format:
    question = answer
    """
    with open(file_path,"r",encoding="utf-8") as f:
        k = f.readline()
    return k



# Configure Gemini API with your API key
genai.configure(api_key=sec.key)  # Replace with your actual Gemini API key

def interact_with_gemini(message):
    """
    Interacts with the Gemini chatbot API.
    Sends a message to Gemini and gets the response.
    """
    cum= load_user_data("user_datax.txt")
    message = cum + message
    # print(message)
    try:
        # Specify the Gemini model you want to use (e.g., "gemini-1.5-flash")
        model = genai.GenerativeModel("gemini-1.5-flash")
        
        # Send the message to the Gemini API for response generation
        
        response = model.generate_content(message)
        
        # Return the response text
        return response.text.strip()
    except Exception as e:
        print(f"Error interacting with Gemini: {e}")
        return None

# Main loop for interacting with Gemini
def main():
    print("Welcome to the Gemini chatbot!")
    while True:
        # Take user input
        user_input = input("Ask Gemini: ")
        if user_input.lower() in ["exit", "quit", "exit chat"]:
            print("Exiting the chat...")
            break

        # Send the input to Gemini and get the response
        gemini_response = interact_with_gemini(user_input)

        # Display the response from Gemini
        if gemini_response:
            print("Gemini says:", gemini_response)
        else:
            print("Sorry, something went wrong with the interaction.")

if __name__ == "__main__":
    # user_data = load_user_data('user_data.txt')
    # print(user_data)
    main()
