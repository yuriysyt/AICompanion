# AI Engine Setup Guide

This document describes how to set up and run the Python AI engine that powers the AI Companion application.

## Overview

The AI engine runs as a separate Python process and communicates with the C# desktop application via gRPC. It uses the IBM Granite 3.3-2B-Instruct model for natural language understanding and command interpretation.

## System Requirements

### Hardware Requirements

For optimal performance with the IBM Granite model:

- **CPU**: Intel Core i5 or AMD Ryzen 5 (8th gen or newer)
- **RAM**: 16 GB minimum (32 GB recommended)
- **GPU**: NVIDIA GPU with 8GB VRAM (optional, for faster inference)
- **Storage**: 10 GB free space for model weights

### Software Requirements

- Python 3.11 or newer
- CUDA 11.8 or newer (for GPU acceleration)
- Windows 10/11 64-bit

## Installation Steps

### Step 1: Create Python Environment

```bash
python -m venv ai_engine_env
ai_engine_env\Scripts\activate
```

### Step 2: Install Dependencies

```bash
pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118
pip install transformers accelerate grpcio grpcio-tools
pip install chromadb sentence-transformers
```

### Step 3: Download IBM Granite Model

```bash
python -c "from transformers import AutoModelForCausalLM; AutoModelForCausalLM.from_pretrained('ibm-granite/granite-3.3-2b-instruct')"
```

The model will be downloaded to your Hugging Face cache directory.

### Step 4: Generate gRPC Code

```bash
python -m grpc_tools.protoc -I../src/AICompanion.Desktop/Protos --python_out=. --grpc_python_out=. ai_service.proto
```

### Step 5: Run the Engine

```bash
python ai_engine_server.py
```

The server will start on localhost:50051.

## Configuration

Environment variables for customization:

| Variable | Default | Description |
|----------|---------|-------------|
| AI_ENGINE_PORT | 50051 | gRPC server port |
| AI_MODEL_PATH | auto | Custom model path |
| AI_USE_GPU | true | Enable GPU acceleration |
| AI_MAX_TOKENS | 512 | Maximum response length |

## Troubleshooting

### Model Loading Fails

If the model fails to load due to memory constraints, try enabling INT4 quantization:

```python
model = AutoModelForCausalLM.from_pretrained(
    "ibm-granite/granite-3.3-2b-instruct",
    load_in_4bit=True,
    device_map="auto"
)
```

### gRPC Connection Refused

Ensure the AI engine is running before starting the desktop application. Check that port 50051 is not blocked by Windows Firewall.

### Slow Inference

For faster inference on CPU, consider using ONNX Runtime optimization or reducing the maximum token length.

## Model Information

The IBM Granite 3.3-2B-Instruct model is designed for instruction following and conversational AI tasks. It is released under the Apache 2.0 license, making it suitable for commercial use.

Key specifications:
- Parameters: 2 billion
- Context length: 8192 tokens
- Training: Instruction-tuned on diverse datasets
- License: Apache 2.0

For more information, visit: https://huggingface.co/ibm-granite/granite-3.3-2b-instruct
