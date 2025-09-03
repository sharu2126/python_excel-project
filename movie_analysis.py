import pandas as pd
import matplotlib.pyplot as plt

# Load dataset

df = pd.read_excel(r"D:\YSM Data Analytics\python_excel project\movie_ratings_reviews.xlsx", 
                   sheet_name="Reviews", engine="openpyxl")

print("File loaded:", df.shape)


# Cleaning
df.columns = df.columns.str.strip().str.lower()
df["reviewcomment"] = df["reviewcomment"].fillna("No comment")
for c in ["moviename","reviewer","reviewcomment"]:
    df[c] = df[c].astype(str).str.strip()
df["rating"] = df["rating"].astype(float)
df["reviewdate"] = pd.to_datetime(df["reviewdate"], errors="coerce")
df.drop_duplicates(inplace=True)

# Analysis
print("Average rating per movie:\n", df.groupby("moviename")["rating"].mean())
print("\nTop 5 reviews:\n", df.sort_values(by="rating", ascending=False).head(5))

# Visualization
df["rating"].plot(kind="hist", bins=10)
plt.title("Movie Ratings")
plt.show()
