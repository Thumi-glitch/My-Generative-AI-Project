class Person:
    def __init__(self, name):
        self.name = name

class Chat:
    def __init__(self, sender, message):
        self.sender = sender
        self.message = message

# Example usage
if __name__ == "__main__":
    player = Person("Alice")
    npc = Person("Bob")
    chat1 = Chat(player.name, "Hello!")
    chat2 = Chat(npc.name, "Hi there!")
    print(f"{chat1.sender}: {chat1.message}")
    print(f"{chat2.sender}: {chat2.message}")
