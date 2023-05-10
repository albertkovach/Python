import math

# define the radius of the circle
radius = 10

# define the symbol to use for the circle
symbol = "*"

# iterate over the rows and columns and check if the point is inside the circle
for row in range(-radius, radius+1):
    for col in range(-radius, radius+1):
        # calculate the distance from the center to the point
        distance = math.sqrt(row**2 + col**2)
        
        # check if the point is inside the circle
        if distance <= radius:
            print(symbol, end="")
        else:
            print(" ", end="")
    print()