class Vehicle:
    wheels = 4

    def __init__(self, name, doors, seats):
        self.name = name
        self.doors = doors
        self.seats = seats

car = Vehicle("Car", 2, 2)
van = Vehicle("van", 6, 12)

print(car.name, car.doors, car.seats)
print(van.name, van.doors, van.seats)