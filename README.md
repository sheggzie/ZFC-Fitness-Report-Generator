# 💪 ZFC Fitness Report Generator

This is a Python-based automation tool designed to generate personalized monthly fitness reports for members of the ZFC (Zeal Fitness Community). 
It collects and processes workout data shared in a WhatsApp group and visualizes individual stats in a clean, shareable image format.

## 📌 Features

* 📊 Aggregates daily workout data (e.g., pushups, squats, planks) submitted by members
* 📅 Automatically compiles monthly summaries per user
* 🖼️ Generates easy-to-read fitness report images
* 📁 Outputs are saved to individual folders, ready for sharing
* 🔧 Customizable for other group fitness communities

## 📂 Folder Structure

```
ZFC-Fitness-Report-Generator/
│
├── data/                  # Raw data files (WhatsApp exports, cleaned sheets, etc.)
├── reports/               # Output image reports
├── src/                   # Core logic and generation scripts
│   └── report_generator.py
├── utils/                 # Helper functions (data cleaning, formatting)
├── templates/             # Image background and style templates
├── README.md              # Project documentation
└── requirements.txt       # Required Python packages
```

## 🛠️ Requirements

* Python 3.8+
* Libraries:

  * `pandas`
  * `matplotlib`
  * `Pillow`
    *(Full list in `requirements.txt`)*

Install with:

```bash
pip install -r requirements.txt
```

## 🚀 Usage

1. **Prepare your data:**

   * Export WhatsApp messages containing daily fitness stats.
   * Clean and save them as `.csv` or `.xlsx` in the `data/` folder.

2. **Run the generator:**

```bash
python src/report_generator.py
```

3. **Find the monthly reports in the `reports/` folder**, each member has their own folder containing an image for the month.

## 🧠 How It Works

* Parses structured daily workout messages (e.g., `Day 1, pushups: 20 reps, 3 sets, plank: 50 secs`)
* Tracks stats like total reps, best performance, and consistency
* Visualizes key metrics in a shareable image

## 🧰 Customization

Want to adapt this for your own community?

* Modify the `src/report_generator.py` script to reflect your custom workout formats
* Replace assets in the `templates/` folder to suit your brand or community

## 🙌 Contributions

Pull requests are welcome! If you have suggestions for improvements or want to add new features, feel free to fork and open a PR.


Made by [@sheggzie](https://github.com/sheggzie)
For questions or support, reach out via [GitHub Issues](https://github.com/sheggzie/ZFC-Fitness-Report-Generator/issues)
