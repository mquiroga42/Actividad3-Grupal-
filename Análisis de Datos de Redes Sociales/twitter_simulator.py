class TwitterSimulator:
    def __init__(self, usuario):
        """
        Initializes the TwitterSimulator object with a given user.
        Args:
            usuario (str): The username for the Twitter simulator.
        """
        self.usuario = usuario
        self.twits = []

    def twittear(self, mensaje):
        """
        Simulates tweeting a message and adds it to the user's timeline.
        Args:
            mensaje (str): The message to be tweeted.
        """
        self.twits.append(mensaje)
        print(f"{self.usuario} has tweeted: {mensaje}")

    def ver_timeline(self):
        """
        Displays the user's timeline with indexed tweets.
        """
        print(f"\n{self.usuario}'s Timeline:")
        for idx, twit in enumerate(self.twits, start=1):
            print(f"{idx}. {twit}")

    def buscar_hashtags(self):
        """
        Searches and prints the top 3 hashtags based on tweet frequency.
        """
        hashtags = {}
        for twit in self.twits:
            words = twit.split()
            for word in words:
                if word.startswith("#"):
                    hashtag = word[1:]
                    hashtags[hashtag] = hashtags.get(hashtag, 0) + 1

        sorted_hashtags = sorted(hashtags.items(), key=lambda x: x[1], reverse=True)

        print("\nAll hashtags found:")
        if not sorted_hashtags:
            print("No hashtags found.")
        else:
            for idx, (hashtag, count) in enumerate(sorted_hashtags[:3], start=1):
                print(f"{idx}. #{hashtag} - {count} mentions (Top {idx})")

# Example of using the Twitter simulator
if __name__ == "__main__":
    usuario1 = "User1"
    twitter_simulator = TwitterSimulator(usuario1)

    # Tweet messages
    twitter_simulator.twittear("Hello, Twitter!")
    twitter_simulator.twittear("Simulating a basic Twitter system.")
    twitter_simulator.twittear("Python is amazing! #Python #Simulation")
    twitter_simulator.twittear("Enjoying the world of programming. #Python #Coding")
    twitter_simulator.twittear("Winter is here, and it's chilly! #Winter")
    twitter_simulator.twittear("Just another winter day. #Winter")
    twitter_simulator.twittear("Programming is my favorite hobby. #Program #Coding")
    twitter_simulator.twittear("Winter vibes and coding. #Winter #Programming")
    twitter_simulator.twittear("Embracing the winter season. #Winter")
    twitter_simulator.twittear("Winter wonderland outside! #Winter")
    twitter_simulator.twittear("Python is the language of choice. #Python #Coding")
    twitter_simulator.twittear("Coding away in the winter chill. #Program #Winter")
    twitter_simulator.twittear("Excited for the summer! #Summer")
    twitter_simulator.twittear("Sunny days ahead. #Summer")
    twitter_simulator.twittear("Chilling by the fireplace. #Winter")
    twitter_simulator.twittear("Snowy days are perfect for coding. #Program #Winter")
    twitter_simulator.twittear("Summer vibes and coding adventures. #Python #Summer")
    twitter_simulator.twittear("Coding marathon in the winter storm. #Program #Winter")
    twitter_simulator.twittear("Winter nights and code snippets. #Program #Winter")
    twitter_simulator.twittear("The joy of coding never ends. #Python #Program #Winter")

    # Search for all hashtags present in the tweets
    twitter_simulator.buscar_hashtags()
