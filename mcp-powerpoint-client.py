import os
import asyncio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from google import genai
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

GMAIL_CREDS_PATH = os.getenv("GMAIL_CREDS_PATH", "credentials.json")
GMAIL_TOKEN_PATH = os.getenv("GMAIL_TOKEN_PATH", "token.json")
api_key = os.getenv("API_KEY")
if not api_key:
    raise ValueError("API_KEY not found in environment variables")

client = genai.Client(api_key=api_key)
# Initialize variables to keep track of the conversation's history.
last_response = None
iteration_responses = []
# Define the maximum number of times the AI can try to solve the task.
max_Iterations = 15
# Track executed tool calls to avoid duplicates
executed_calls = set()

def reset_state():
    global last_response, iteration_responses, executed_calls
    last_response = None
    iteration_responses = []
    executed_calls = set()

async def generate_with_timeout(client, prompt, timeout=30):
    print("[LLM] Starting LLM generation...")
    try:
        loop = asyncio.get_event_loop()
        response = await asyncio.wait_for(
            loop.run_in_executor(
                None,
                lambda: client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=prompt
                )
            ),
            timeout=timeout
        )
        print("[LLM] LLM generation completed")
        return response
    except TimeoutError:
        print("[LLM] LLM generation timed out!")
        raise
    except Exception as e:
        print(f"[LLM] Error in LLM generation: {e}")
        raise


async def main():
    # --- 1. Preparation and Server Setup ---

    # Reset any previous state or data, ensuring a clean start.
    reset_state()

    print("[CLIENT] Configuring server parameters...")
    # Define how to start the PowerPoint 'server' (a separate program).
    # StdioServerParameters helps define this external process.
    ppt_server_params = StdioServerParameters(
        command="uv",
        args=["run", "mcp-server.py"],  # Added --no-project flag
        cwd=os.getcwd()  # Explicitly set working directory
    )
    print(f"[CLIENT] PowerPoint server command: uv run mcp-server.py")

    # Define how to start the Gmail 'server' (a separate program).
    # It needs extra arguments for credentials and token paths
    gmail_server_params = StdioServerParameters(
        command="uv",
        args=[
            "run",  "gmail-server.py",  # Added --no-project flag
            "--creds-file-path", os.path.abspath(GMAIL_CREDS_PATH),  # Use absolute path
            "--token-path", os.path.abspath(GMAIL_TOKEN_PATH)       # Use absolute path
        ],
        cwd=os.getcwd()  # Explicitly set working directory
    )
    
    print("[CLIENT] Launching both MCP servers (PowerPoint & Gmail)...")

    # --- 2. Starting Servers and Establishing Connections ---
    try:
        print("[CLIENT] Starting PowerPoint server...")
        await asyncio.sleep(1)
        
        async with stdio_client(ppt_server_params) as (ppt_read, ppt_write):
            try:
                async with ClientSession(ppt_read, ppt_write) as ppt_session:
                    print("[CLIENT] Initializing PowerPoint session...")
                    try:
                        await ppt_session.initialize()
                        print("[CLIENT] PowerPoint server connected")
                    except Exception as e:
                        print(f"[ERROR] Failed to initialize PowerPoint session: {str(e)}")
                        raise

                    print("[CLIENT] Starting Gmail server...")
                    await asyncio.sleep(1)
                    
                    async with stdio_client(gmail_server_params) as (gmail_read, gmail_write):
                        try:
                            async with ClientSession(gmail_read, gmail_write) as gmail_session:
                                print("[CLIENT] Initializing Gmail session...")
                                try:
                                    await gmail_session.initialize()
                                    print("[CLIENT] Gmail server connected")
                                    print("[CLIENT] Both servers are ready")

                                    # --- 3. Discovering Available Tools ---
                                    print("[CLIENT] Fetching tool lists from both servers...")
                                    ppt_tools_result = await ppt_session.list_tools()
                                    gmail_tools_result = await gmail_session.list_tools()

                                    ppt_tools = ppt_tools_result.tools
                                    gmail_tools = gmail_tools_result.tools

                                    print(f"[CLIENT] PowerPoint tools: {[t.name for t in ppt_tools]}")
                                    print(f"[CLIENT] Gmail tools: {[t.name for t in gmail_tools]}")

                                    # --- 4. Preparing Tools for the AI (LLM) ---
                                    tool_to_session = {}
                                    all_tools = []

                                    for tool in ppt_tools:
                                        tool_to_session[tool.name] = ppt_session
                                        all_tools.append(tool)

                                    for tool in gmail_tools:
                                        tool_to_session[tool.name] = gmail_session
                                        all_tools.append(tool)

                                    tools_description = []
                                    for tool in all_tools:
                                        name = tool.name
                                        desc = getattr(tool, "description", "No description available")
                                        params = tool.inputSchema
                                        if "properties" in params:
                                            param_details = []
                                            for pname, pinfo in params["properties"].items():
                                                ptype = pinfo.get("type", "unknown")
                                                param_details.append(f"{pname}: {ptype}")
                                            params_str = ", ".join(param_details)
                                        else:
                                            params_str = "no parameters"
                                        #tools_description.append(f"{name}({params_str}) - {desc}")
                                         #Fix: remove () from tool names in description to avoid LLM confusion
                                        tools_description.append(f"{name} - {desc}\nArgs: {params_str}")

                                    tools_description = "\n".join(tools_description)
                                    print("[CLIENT] Merged tool descriptions for LLM prompt:")
                                    print(tools_description)

                                    # --- 5. Setting up the Conversation with the AI (LLM) ---
                                    system_prompt = f"""
You are an automation agent capable of controlling PowerPoint and Gmail through MCP tools. You can create presentations, draw shapes, add content, and send emails through various available tools.

Available Tools:
{tools_description}

You must respond with EXACTLY ONE line in one of these formats (no additional text):
1. For function calls:
   FUNCTION_CALL: function_name|param1|param2|...
   
2. For final answers when task is complete:
   FINAL_ANSWER: [message]

Important Rules:
- Process one action at a time in the correct sequence
- Wait for each tool's response before proceeding to next action
- Do not repeat the same FUNCTION_CALL with identical parameters once it has been executed and succeeded
- Always check the conversation history: if a tool call already appears there, do not issue it again
- If the previous tool call succeeded, move on to the next step instead of repeating it
- If the user asks to send the presentation as an attachment, always pass the FULL file path returned by save_presentation in the TOOL_RESULT as the attachment_path argument to the email function. 
- Do not invent or shorten the filename. Use exactly the path shown in the TOOL_RESULT of save_presentation.
- Only give FINAL_ANSWER after all steps are completed


Examples:
- FUNCTION_CALL: open_powerpoint
- FUNCTION_CALL: draw_rectangle_with_text|Hello World|100|100|200|100
- FUNCTION_CALL: send-email|user@example.com|Meeting Summary|The presentation is ready
- FINAL_ANSWER: [Task completed successfully]

DO NOT include any explanations or additional text.
Your entire response should be a single line starting with either FUNCTION_CALL: or FINAL_ANSWER:"""

                                    global last_response, iteration_responses, executed_calls
                                    query = input("\n[CLIENT] Enter your command: ")
                                    print(f"[CLIENT] User input: {query}")

                                    # --- 6. The Main Loop (AI Thinking and Tool Execution) ---
                                    for iteration in range(max_Iterations):
                                        print(f"\n[CLIENT] --- Iteration {iteration + 1} ---")

                                        # --- Conversation history handling ---
                                        if last_response is None:
                                            # First time: just send the user query
                                            current_query = query
                                        else:
                                            # Subsequent times: include user query + all tool call results so far
                                            history = "\n".join(iteration_responses)
                                            current_query = f"{query}\n{history}"

                                        prompt = f"{system_prompt}\n\nConversation so far:\n{current_query}"
                                        print("[CLIENT] Sending prompt to LLM:")
                                        print(prompt)

                                        try:
                                            response = await generate_with_timeout(client, prompt)
                                            response_text = response.text.strip()
                                            print(f"[LLM] RAW RESPONSE: {response_text}")

                                            # --- A. The AI Wants to Use a Tool (FUNCTION_CALL) ---
                                            if response_text.startswith("FUNCTION_CALL:"):
                                                _, function_info = response_text.split(":", 1)
                                                parts = [p.strip() for p in function_info.split("|")]
                                                func_name, params = parts[0], parts[1:]
                                                
                                                # Fix: normalize function name (strip trailing parentheses if present)
                                                if func_name.endswith("()"):
                                                    func_name = func_name[:-2]
                                                    
                                                print(f"[CLIENT] LLM selected tool: {func_name} with params: {params}")

                                                tool_session = tool_to_session.get(func_name)
                                                tool = next((t for t in all_tools if t.name == func_name), None)

                                                if not tool or not tool_session:
                                                    print(f"[CLIENT] ERROR: Unknown tool requested: {func_name}")
                                                    break

                                                arguments = {}
                                                schema_properties = tool.inputSchema.get("properties", {})
                                                for pname, pinfo in schema_properties.items():
                                                    if not params:
                                                        break
                                                    raw_value = params.pop(0)
                                                    # Handle key=value style params (normalize)
                                                    if "=" in raw_value:
                                                        _, raw_value = raw_value.split("=", 1)
                                                    ptype = pinfo.get("type", "string")
                                                    if ptype == "integer":
                                                        arguments[pname] = int(raw_value)
                                                    elif ptype == "number":
                                                        arguments[pname] = float(raw_value)
                                                    else:
                                                        arguments[pname] = str(raw_value)

                                                call_signature = f"{func_name}|{arguments}"
                                                if call_signature in executed_calls:
                                                    print(f"[CLIENT] Skipping duplicate tool call: {call_signature}")
                                                    continue
                                                executed_calls.add(call_signature)

                                                print(f"[CLIENT] Invoking tool '{func_name}' on correct server with arguments: {arguments}")
                                                result = await tool_session.call_tool(func_name, arguments=arguments)

                                                if hasattr(result, "content"):
                                                    outputs = [c.text for c in result.content if hasattr(c, "text")]
                                                    print(f"[CLIENT] TOOL RESULT: {outputs}")
                                                    last_response = outputs
                                                    # Stop only if the final tool in workflow succeeds
                                                    if func_name in ["send_email__via_outlook_with_attachment", "send-email"] and any("successfully" in o.lower() for o in outputs):
                                                        print("[CLIENT] Task completed, stopping after final step.")
                                                        break
                                                else:
                                                    last_response = str(result)
                                                    print(f"[CLIENT] TOOL RESULT: {last_response}")

                                                # --- New history update ---
                                                # Store both the function call and result for next loop
                                                iteration_responses.append(f"TOOL_CALL: {func_name} args={arguments}")
                                                iteration_responses.append(f"TOOL_RESULT: {last_response}")

                                            # --- B. The AI Has Finished the Task (FINAL_ANSWER) ---
                                            elif response_text.startswith("FINAL_ANSWER:"):
                                                print(f"[CLIENT] DONE: Final Answer: {response_text}")
                                                break

                                        except Exception as e:
                                            print(f"[CLIENT] EXCEPTION: {e}")
                                            break
                                except Exception as e:
                                    print(f"[ERROR] Failed to initialize Gmail session: {str(e)}")
                                    raise
                        except Exception as e:
                            print(f"[ERROR] Failed to create Gmail session: {str(e)}")
                            raise
            except Exception as e:
                print(f"[ERROR] Failed to create PowerPoint session: {str(e)}")
                raise
    except Exception as e:
        print(f"[ERROR] Fatal error in server initialization: {str(e)}")
        return

if __name__ == "__main__":
    asyncio.run(main())