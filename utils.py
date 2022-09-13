def index_of_clip(clips, currentClip):
    for x in range(0, len(clips)):
        if f"{clips[x]}" == f"{currentClip}":
            return x + 1

    return -1
