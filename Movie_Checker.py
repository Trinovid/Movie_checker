import os
from pymediainfo import MediaInfo
import xlsxwriter

def get_metadata(file_path):
    """Get detailed metadata for the movie file."""
    try:
        media_info = MediaInfo.parse(file_path)
        metadata = {
            "File Size": "Unknown",
            "Resolution": "Unknown",
            "Video Codec": "Unknown",
            "Audio Codec": "Unknown",
            "Audio Channels": "Unknown",
            "Bitrate": "Unknown",
            "HDR": "Unknown",
            "Aspect Ratio": "Unknown",
            "Color Depth": "Unknown"
        }
        # Add file size
        size_bytes = os.path.getsize(file_path)
        size_gb = size_bytes / (1024 ** 3)
        metadata["File Size"] = f"{size_gb:.2f} GB"

        for track in media_info.tracks:
            if track.track_type == "Video":
                metadata["Resolution"] = f"{track.height}p" if track.height else "Unknown"
                metadata["Video Codec"] = track.format
                metadata["Bitrate"] = f"{int(track.bit_rate) // 1000} kbps" if track.bit_rate else "Unknown"
                metadata["Aspect Ratio"] = f"{track.display_aspect_ratio}" if track.display_aspect_ratio else "Unknown"
                metadata["HDR"] = track.transfer_characteristics if track.transfer_characteristics else "Unknown"
                metadata["Color Depth"] = f"{track.bit_depth}-bit" if track.bit_depth else "Unknown"
            elif track.track_type == "Audio":
                metadata["Audio Codec"] = track.format
                metadata["Audio Channels"] = f"{track.channel_s}" if track.channel_s else "Unknown"

        return metadata
    except Exception as e:
        print(f"Error reading metadata for {file_path}: {e}")
        return metadata

def scan_movies_to_excel(directory, output_file):
    """Scan the directory for movie files and save details to an Excel file."""
    movie_extensions = ('.mp4', '.mkv', '.avi', '.mov')  # Add more extensions as needed
    movie_list = []

    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(movie_extensions):
                file_path = os.path.join(root, file)
                metadata = get_metadata(file_path)
                movie_name = os.path.splitext(file)[0]
                movie_list.append([
                    movie_name, file_path, metadata["File Size"], metadata["Resolution"], 
                    metadata["Video Codec"], metadata["Audio Codec"], metadata["Audio Channels"], 
                    metadata["Bitrate"], metadata["HDR"], metadata["Aspect Ratio"], metadata["Color Depth"]
                ])

    # Write to Excel
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet("Movies")

    # Define header format
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    link_format = workbook.add_format({'font_color': 'blue', 'underline': True})

    # Write headers
    headers = [
        "Movie Name", "Location", "File Size", "Resolution", "Video Codec", 
        "Audio Codec", "Audio Channels", "Bitrate", "HDR", "Aspect Ratio", "Color Depth"
    ]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)

    # Write movie data
    for row_num, movie in enumerate(movie_list, start=1):
        for col_num, data in enumerate(movie):
            if col_num == 1:  # Location column
                worksheet.write_url(row_num, col_num, f"file:///{data}", link_format, data)
            else:
                worksheet.write(row_num, col_num, data)

    workbook.close()
    print(f"Movie details saved to {output_file}")

# User-specified directory and output file
movie_directory = input("Enter the directory to scan: ").strip()
output_excel_file = "movie_list.xlsx"

scan_movies_to_excel(movie_directory, output_excel_file)
