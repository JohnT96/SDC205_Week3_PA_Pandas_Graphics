import pandas as pd
import matplotlib.pyplot as plt

# John Tenney
# SDC205 Week 3 Performance Assessment - Pandas Graphics

student_id = "johten7331"

ballrooms = {
    "Names": ["Ballroom 1", "Ballroom 2", "Ballroom 3"],
    "Capacity": [25000, 11000, 5000]
}

demographics = {
    "Children": 18000,
    "Adults": 13000,
    "Teens": 10000
}

df = pd.DataFrame(ballrooms)

print(student_id)
print(df)

df.plot(
    x="Names",
    y="Capacity",
    kind="bar",
    title="Ballroom Capacity"
)

plt.xlabel("Ballrooms")
plt.ylabel("Capacity")
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()

plt.pie(
    demographics.values(),
    labels=demographics.keys(),
    autopct="%1.1f%%"
)

plt.title("Event Demographics")
plt.axis("equal")
plt.show()