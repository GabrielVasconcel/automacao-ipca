import gradio
import safehttpx
import groovy
import os

print("--- COPIE OS CAMINHOS ABAIXO ---")
print(f"GRADIO:    {os.path.dirname(gradio.__file__)}")
print(f"SAFEHTTPX: {os.path.dirname(safehttpx.__file__)}")
print(f"GROOVY:    {os.path.dirname(groovy.__file__)}")