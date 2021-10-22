import pandas as pd
import numpy as np
import pickle
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder, StandardScaler
from sklearn.neighbors import KNeighborsClassifier

df = pd.read_excel('sheets/cash.xls')

df = df[['Description','AcctCode','Invoice Line Total','EscrowBank','TitleCoNum','OrderCategory']]

def clean_data(df):
    df.loc[(df.AcctCode == 40000) & (df.OrderCategory.isin([2,5])), 'AcctCode'] = 40002
    return df

def load_dataframe(file):
    if file.endswith('.csv'):
        return pd.read_csv(file)
    if file.endswith('.xls') or file.endswith('.xlsx'):
        return pd.read_excel(file)
    return None

def save_data(df, save_file):
    with open(save_file, 'wb') as f:
        pickle.dump(df, f)

def load_data(save_file):
    with open(save_file, 'rb') as f:
        return pickle.load(f)

def add_data(df):
    data = load_data('data/dataframe.pkl')
    df = pd.concat([data, df], ignore_index=True)
    save_data(df, 'data/dataframe.pkl')

def train_data(data_file='data/dataframe.pkl', train=False):
    df = load_data(data_file)
    df.dropna(how='any', inplace=True)
    df.loc[(df.AcctCode == 40000) & (df.OrderCategory.isin([2,5])), 'AcctCode'] = 40002

    description_encoder = LabelEncoder()
    oc_encoder = LabelEncoder()
    escrow_encoder = LabelEncoder()
    scaler = StandardScaler()
    
    df.Description = description_encoder.fit_transform(df.Description)
    df.OrderCategory = oc_encoder.fit_transform(df.OrderCategory)
    df.EscrowBank = escrow_encoder.fit_transform(df.EscrowBank)
    df['Invoice Line Total'] = scaler.fit_transform(df['Invoice Line Total'].to_numpy().reshape(-1,1))

    X, y = df[['Description','Invoice Line Total','OrderCategory','EscrowBank']].values, df['AcctCode'].values
    X_train, X_test, y_train, y_test = train_test_split(X, y, random_state=42)

    if train:
        model = None
        score = 0
        for x in [1, 2, 3, 5, 7, 10]:
            print('n_neighbors: {}'.format(x))
            X_train, X_test, y_train, y_test = train_test_split(X, y, random_state=42)
            knn = KNeighborsClassifier(n_neighbors=x)
            knn.fit(X_train, y_train)
            model_score = knn.score(X_test, y_test)
            print('Model Score: {}'.format(model_score))
            if model_score > score:
                score = model_score
                model = knn
            print()
        predictions = model.predict(X_test)
        print(predictions, y_test)
        
        with open('data/acct_predictor.pkl', 'wb') as f:
            pickle.dump(model, f)
    else:
        with open('data/acct_predictor.pkl', 'rb') as f:
            model = pickle.load(f)

        X_train, X_test, y_train, y_test = train_test_split(X, y, random_state=42)
        
    predictions = model.predict(X_test)

    correct, wrong = [], []
    for i in range(len(predictions)):
        if predictions[i] == y_test[i]:
            correct.append('Accurate: {} = {}'.format(predictions[i], y_test[i]))
        else:
            wrong.append('Wrong: {} = {}'.format(predictions[i], y_test[i]))

    print(wrong)


if __name__ == "__main__":
    train_data(train=True)