.PHONY: install run agent agent-schedule app

install:
	uv sync
	uv run python -m playwright install chromium

run:
	uv run python bid_all.py

agent:
	uv run python agent.py

agent-schedule:
	uv run python agent.py --schedule

app:
	uv run streamlit run app.py
