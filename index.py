import pandas as pd

class ABCAnalysis:
    def __init__(self, filename):
        self.data = pd.read_excel(filename)

    def calculate_ABC(self):
        print('total entries - ', len(self.data))
        # Calculate total value of products (Quantity X Price)
        self.data['total_value'] = self.data['total_inventory'] * self.data['price_per_unit']

        # Group data by product code and sum total values
        # to take of the case where a product has multiple entries
        product_total_value = self.data.groupby('product_code')['total_value'].sum().reset_index()
        print('\nunique product count - ', len(product_total_value))

        # Sort products by their values
        sorted_data = product_total_value.sort_values(by='total_value', ascending=False)
        # print('\nsorted data', sorted_data)
        
        # Calculate the thresholds
        total_value = sorted_data['total_value'].sum()
        A_threshold = 0.8 * total_value
        B_threshold = 0.95 * total_value

        # classification of products into their classes
        A_products = []
        B_products = []
        C_products = []
        current_value = 0
        for index, row in sorted_data.iterrows():
            # print('\n', row)
            current_value += row['total_value']
            if current_value <= A_threshold:
                A_products.append({'product_code': row['product_code'], 'total_value': row['total_value']})
            elif current_value <= B_threshold:
                B_products.append({'product_code': row['product_code'], 'total_value': row['total_value']})
            else:
                C_products.append({'product_code': row['product_code'], 'total_value': row['total_value']})

        return A_products, B_products, C_products

abc_analysis = ABCAnalysis('./inventory data.xlsx')
A_products, B_products, C_products = abc_analysis.calculate_ABC()

# write to excel
columns = ['product_code', 'total_value']
A_products = pd.DataFrame(A_products, columns=columns)
A_products.to_excel('A_products.xlsx', index=False)

B_products = pd.DataFrame(B_products, columns=columns)
B_products.to_excel('B_products.xlsx', index=False)

C_products = pd.DataFrame(C_products, columns=columns)
C_products.to_excel('C_products.xlsx', index=False)

print('Done writing to files')
# print("Products A's:", A_products)
# print("Products A's:", len(A_products))
# print("Products B's:", B_products)
# print("Products B's:", len(B_products))
# print("Products C's:", C_products)
# print("Products C's:", len(C_products))