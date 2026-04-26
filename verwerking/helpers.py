# verwerking/helpers.py

maanden = [
    "September", "Oktober", "November", "December",
    "Januari", "Februari", "Maart", "April", "Mei", "Juni"
]

maanden_kort = [m[:3] for m in maanden]

def get_jaar(maand: str, jaar: str) -> str:
    # Sep–Dec in 'jaar', Jan–Jun in 'jaar+1'
    return jaar if maanden.index(maand) < 4 else str(int(jaar) + 1)
