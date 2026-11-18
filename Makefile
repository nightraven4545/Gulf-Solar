MODEL := model/NREP_Tariff_LCOE_Model.xlsx

.PHONY: data model deck check all

data:
	uv run python -m nrep.fetch

# openpyxl writes formulas only; LibreOffice recalculates so the saved file carries values.
model:
	uv run python -m nrep.build_model
	soffice --headless --calc --convert-to xlsx --outdir build $(MODEL) >/dev/null
	mv build/$(notdir $(MODEL)) $(MODEL)

deck: model
	uv run python -m nrep.build_deck

check:
	uv run ruff check src tests
	uv run pytest -q

all: model deck check
