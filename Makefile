.PHONY: install run agent agent-schedule

install:
	uv sync
	uv run playwright install chromium

run:
	uv run python bid_all.py

agent:
	uv run python agent.py

agent-schedule:
	uv run python agent.py --schedule
