import pandas as pd
import torch
from transformers import BertTokenizer, BertForSequenceClassification, get_linear_schedule_with_warmup
from torch.utils.data import DataLoader, RandomSampler, SequentialSampler, TensorDataset
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import accuracy_score
from torch.nn.utils.rnn import pad_sequence
from torch.optim import AdamW

# Verify library versions
import transformers
print(f"Transformers version: {transformers.__version__}")
print(f"Torch version: {torch.__version__}")

# Load the data
file_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Master Expense Report V4.0.xlsx"
df = pd.read_excel(file_path, sheet_name='Fixed', usecols=['New', 'Notes'])

# Fill NaN values in 'Notes' column with empty strings and drop rows with NaN in 'New' column
df['Notes'] = df['Notes'].fillna('')
df = df.dropna(subset=['New'])

# Load the BERT tokenizer
tokenizer = BertTokenizer.from_pretrained('bert-base-uncased')

# Tokenize the 'Notes' column
df['tokenized_notes'] = df['Notes'].apply(lambda x: tokenizer.encode(x, add_special_tokens=True) if x.strip() != "" else [101, 102])

# Define the maximum sequence length for BERT
max_length = 512

# Pad the tokenized sequences using PyTorch's pad_sequence
tokenized_notes = [torch.tensor(note) for note in df['tokenized_notes'].values]
input_ids = pad_sequence(tokenized_notes, batch_first=True, padding_value=tokenizer.pad_token_id)
input_ids = input_ids[:, :max_length]  # Truncate sequences to max_length if necessary

# Ensure input_ids are of type LongTensor
input_ids = input_ids.long()

# Create attention masks
attention_masks = (input_ids != tokenizer.pad_token_id).long()

# Extract the target variable
label_encoder = LabelEncoder()
labels = label_encoder.fit_transform(df['New'].values)
labels = torch.tensor(labels).long()

# Split the data into training and testing sets
input_ids_train, input_ids_test, attention_masks_train, attention_masks_test, labels_train, labels_test = train_test_split(
    input_ids, attention_masks, labels, test_size=0.3, random_state=42
)

# Create DataLoader for training and testing sets
train_data = TensorDataset(input_ids_train, attention_masks_train, labels_train)
train_sampler = RandomSampler(train_data)
train_dataloader = DataLoader(train_data, sampler=train_sampler, batch_size=8)

test_data = TensorDataset(input_ids_test, attention_masks_test, labels_test)
test_sampler = SequentialSampler(test_data)
test_dataloader = DataLoader(test_data, sampler=test_sampler, batch_size=8)

# Load a pre-trained BERT model for sequence classification
model = BertForSequenceClassification.from_pretrained('bert-base-uncased', num_labels=len(label_encoder.classes_))

# Define optimizer and learning rate scheduler
optimizer = AdamW(model.parameters(), lr=2e-5, eps=1e-8)
total_steps = len(train_dataloader) * 16  # Number of training steps (epochs * batches)
scheduler = get_linear_schedule_with_warmup(optimizer, num_warmup_steps=0, num_training_steps=total_steps)

# Move model to GPU if available
device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
model.to(device)

# Training loop
epochs = 3
for epoch in range(epochs):
    model.train()
    total_loss = 0
    for step, batch in enumerate(train_dataloader):
        batch_inputs, batch_masks, batch_labels = batch
        batch_inputs, batch_masks, batch_labels = batch_inputs.to(device), batch_masks.to(device), batch_labels.to(device)
        
        model.zero_grad()
        outputs = model(batch_inputs, attention_mask=batch_masks, labels=batch_labels)
        loss = outputs.loss
        total_loss += loss.item()
        loss.backward()
        optimizer.step()
        scheduler.step()
    
    avg_train_loss = total_loss / len(train_dataloader)
    print(f"Epoch {epoch + 1}, Average Training Loss: {avg_train_loss}")

# Evaluation loop
model.eval()
predictions, true_labels = [], []
for batch in test_dataloader:
    batch_inputs, batch_masks, batch_labels = batch
    batch_inputs, batch_masks, batch_labels = batch_inputs.to(device), batch_masks.to(device), batch_labels.to(device)
    
    with torch.no_grad():
        outputs = model(batch_inputs, attention_mask=batch_masks)
        logits = outputs.logits
    
    predictions.extend(torch.argmax(logits, dim=1).tolist())
    true_labels.extend(batch_labels.tolist())

# Calculate accuracy
accuracy = accuracy_score(true_labels, predictions)
print(f"Evaluation Accuracy: {accuracy}")
