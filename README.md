# Excel Form Web Application

This Node.js web application allows users to submit data through a form, and the submitted data is stored in an Excel file. Additionally, users can view the data from the Excel file on the webpage.

## Getting Started

Follow these steps to run the application locally:

### Prerequisites

- Node.js installed on your machine
- Git installed on your machine

### Installation

1. Clone the repository:

   ```bash
   git clone <repository-url>
   ```

2. Navigate to the project directory:

   ```bash
   cd <project-directory>
   ```

3. Install dependencies:

   ```bash
   npm install
   ```

### Running the Application

1. Start the Node.js server:

   ```bash
   npm start
   ```

2. Open your web browser and go to [http://localhost:3000](http://localhost:3000).

## Usage

1. Access the web page and fill out the Excel form with your name and age.
2. Click the "Submit" button to save the data to the Excel file.
3. View the submitted data on the webpage in a tabular format.

## Project Structure

- **public/index.html**: HTML file containing the structure of the webpage.
- **public/style.css**: CSS file for styling the webpage.
- **public/script.js**: JavaScript file for fetching and displaying data on the webpage.
- **server.js**: Node.js server file with Express configuration.
- **data.xlsx**: Excel file to store submitted data.

## Dependencies

- [Express](https://expressjs.com/): Web framework for Node.js.
- [xlsx](https://www.npmjs.com/package/xlsx): Library for reading and writing Excel files.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with Node.js, Express, and xlsx.
- Inspired by the need for a simple web application to manage Excel data.

Feel free to contribute or open issues if you have any suggestions or improvements!
