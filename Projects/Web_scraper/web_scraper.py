from requests_html import HTMLSession  # Import HTMLSession from the requests_html module to make HTTP requests and parse HTML.
import csv  # Import the csv module to handle CSV file operations.

# Create a session object for making HTTP requests.
s = HTMLSession()

# Function to get product links from a specific page.
def get_product_links(page):
    # Construct the URL for the given page number.
    url = f'https://themes.woocommerce.com/storefront/product-category/clothing/page/{page}'
    links = []  # List to store product links.
    
    # Send a GET request to the URL and get the response.
    response = s.get(url)
    
    # Find all product elements on the page.
    products = response.html.find('ul.products li')
    
    # Loop through each product item and extract the product link.
    for item in products:
        links.append(item.find('a', first=True).attrs['href'])
    
    # Return the list of product links.
    return links

# Function to parse product details from a given product URL.
def parse_product(url):
    # Send a GET request to the product URL and get the response.
    response = s.get(url)
    
    # Extract the product title.
    title = response.html.find('h1.product_title.entry-title', first=True).text.strip()
    
    # Try to extract the product description; if not found, set description to 'None'.
    try:
        description = response.html.find('div.woocommerce-product-details__short-description', first=True).text.strip()
    except AttributeError as err:
        description = 'None'
    
    # Extract the product price, removing any newline characters.
    price = response.html.find('p.price', first=True).text.strip().replace('\n', '')
    
    # Extract the product category.
    category = response.html.find('span.posted_in a', first=True).text.strip()
    
    # Try to extract the SKU (Stock Keeping Unit); if not found, set SKU to 'None'.
    try:
        sku = response.html.find('span.sku', first=True).text.strip()
    except AttributeError as err:
        sku = 'None'

    # Create a dictionary to store the product details.
    product = {
        'title': title,
        'description': description,
        'price': price,
        'sku': sku,
        'category': category,
    }
    
    # Return the product details dictionary.
    return product

# Function to save the scraped product details to a CSV file.
def save_csv(results):
    # Get the dictionary keys (column names) from the first product.
    keys = results[0].keys()
    
    # Open a CSV file for writing.
    with open('scraped_data.csv', 'w') as f:
        # Create a CSV DictWriter object with the keys as the field names.
        dict_writer = csv.DictWriter(f, keys)
        
        # Write the header (column names) to the CSV file.
        dict_writer.writeheader()
        
        # Write all product details (rows) to the CSV file.
        dict_writer.writerows(results)

# Main function to coordinate the scraping process.
def main():
    results = []  # List to store all the scraped product details.
    
    # Loop through pages 1 to 3.
    for x in range(1, 4):
        print('Getting Page ', x)
        
        # Get all product links from the current page.
        urls = get_product_links(x)
        
        # Loop through each product link and parse the product details.
        for url in urls:
            results.append(parse_product(url))
        
        # Print the total number of products scraped so far.
        print('Total results: ', len(results))
        
        # Save the current results to the CSV file.
        save_csv(results)

# Check if the script is being run directly (not imported as a module).
if __name__ == '__main__':
    main()  # Call the main function to start the scraping process.
