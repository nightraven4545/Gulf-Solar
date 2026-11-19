MODEL := model/NREP_Tariff_LCOE_Model.xlsx

.PHONY: data model deck check all

data:
	uv run python -m nrep.fetch

# openpyxl writes formulas only; LibreOffice recalculates so the saved file carries values.
model:
	uv run python -m nrep.build_model
	mkdir -p build
	soffice --headless --calc --convert-to xlsx --outdir build $(MODEL) >/dev/null
	mv build/$(notdir $(MODEL)) $(MODEL)

deck:
	uv run python -m nrep.build_map
	uv run python -m nrep.build_deck
	soffice --headless --convert-to pdf --outdir deck deck/NREP_Tariff_Replicability.pptx >/dev/null
	soffice --headless --convert-to pdf --outdir onepager onepager/NREP_onepager.pptx >/dev/null
	pdftoppm -f 15 -l 15 -r 110 -png -singlefile deck/NREP_Tariff_Replicability.pdf docs/waterfall

check:
	uv run ruff check src tests
	uv run pytest -q

all: model deck check
