from fastapi import FastAPI
from pydantic import BaseModel
from lib import banners, goreport
import os
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()
allowed_origins = ["*"]

# Configure CORS settings
app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class Campagne(BaseModel):
    id: str


def parse_options(self, id, format, combine, complete, config, google, verbose):
    """GoReport uses the Gophish API to connect to your Gophish instance using the
    IP address, port, and API key for your installation. This information is provided
    in the gophish.config file and loaded at runtime. GoReport will collect details
    for the specified campaign and output statistics and interesting data for you.

    Select campaign ID(s) to target and then select a report format.\n
       * csv: A comma separated file. Good for copy/pasting into other documents.\n
       * word: A formatted docx file. A template.docx file is required (see the README).\n
       * quick: Command line output of some basic stats. Good for a quick check or client call.\n
    """
    # Print the Gophish banner
    banners.print_banner()
    # Create a new Goreport object that will use the specified report format
    gophish = goreport.Goreport("word", config, google, verbose)
    # Execute reporting for the provided list of IDs
    gophish.run(id, combine, complete)


@app.get("/")
def read_root():
    return {"message": "Hello, FastAPI!"}


@app.get("/generateReport/{id}")
async def create_item(id: str):
    print(id)
    # banners.print_banner()
    gophish = goreport.Goreport("word", "./Gophish.config", "", "")
    print(gophish)
    gophish.run(id, False, False)
    for filename in os.listdir("./"):
        if filename.startswith("Gophish Results for "+ gophish.safe_name):
            # Found a file whose name starts with the specified word
            full_file_path = os.path.join("./", filename)
            print(f"Found file: {full_file_path}")
            return FileResponse(full_file_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="Gophish Results for "+ gophish.safe_name + ".docx")

