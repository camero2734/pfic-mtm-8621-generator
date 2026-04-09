"""

AAB = adjusted adjusted basis; UNI = unreversed inclusions.

§1296 requires:
  - MTM inclusion = FMV(EOY) - AAB(BOY), not FMV(EOY) - FMV(BOY)
  - AAB rolls forward: AAB(EOY) = AAB(BOY) + gain - allowed_loss
  - UNI rolls forward: UNI(EOY) = UNI(BOY) + gain - allowed_loss
  - Loss limitation is per-lot
"""

import numpy as np
import pandas as pd
import pytest

from main import compute_lot, compute_part1, LotResult


FILEPATH = "test_input.xlsx"


def _dt(year, month=1, day=1):
    return pd.Timestamp(year=year, month=month, day=day)


def _make_eoy_df(rows):
    return pd.DataFrame(rows, columns=["Year", "Price", "Exchange Rate"])


def _make_lot_df(rows):
    defaults = {
        "Date: Acquisition": None,
        "Price per share: Acquisition": None,
        "Number of shares": None,
        "Cost: Acquisition": None,
        "Exchange Rate: Acquisition": None,
        "Date: Sale": None,
        "Price per share: Sale": np.nan,
        "Exchange Rate: Sale": np.nan,
    }
    result = []
    for row in rows:
        result.append({**defaults, **row})
    return pd.DataFrame(result)


# -- Acquisition year (AAB = original basis) --


class TestAcquisitionYearHolding:
    def test_gain_in_acquisition_year(self):
        # 100 shares @ $50, FX 1.0; EOY price $60 → gain $1000
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 60.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.adjusted_basis == 5000
        assert result.fmv == 6000
        assert result.gain_loss == 1000

    def test_loss_in_acquisition_year_no_uni(self):
        # Same lot; EOY price $40 → non-deductible loss (no UNI)
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 40.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.gain_loss == -1000
        assert result.adjusted_basis == 5000
        assert result.unreversed == 0
        assert result.ordinary_loss == 0


# -- Second year after a gain (AAB = FMV prior year, coincidentally correct) --


class TestSecondYearGainAfterGain:
    def test_gain_after_prior_gain(self):
        # Year 1: gain → AAB(EOY) = FMV. Year 2: another gain on that basis.
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2022),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2022, 60.0, 1.0), (2023, 75.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.adjusted_basis == 6000
        assert result.fmv == 7500
        assert result.gain_loss == 1500


# -- Sale scenarios --


class TestSaleInCurrentYear:
    def test_sale_gain(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                    "Date: Sale": _dt(2023, 6, 15),
                    "Price per share: Sale": 65.0,
                    "Exchange Rate: Sale": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 60.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.is_holding is False
        assert result.proceeds == 6500
        assert result.adjusted_basis == 5000
        assert result.sale_gain_loss == 1500

    def test_sale_loss_with_uni(self):
        # Sale loss with AAB > original basis → ordinary loss limited by UNI
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2022),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                    "Date: Sale": _dt(2023, 6, 15),
                    "Price per share: Sale": 55.0,
                    "Exchange Rate: Sale": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2022, 70.0, 1.0), (2023, 60.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.sale_gain_loss == -1500
        assert result.adjusted_basis == 7000
        assert result.unreversed == 2000
        assert result.ordinary_loss == -1500

    def test_sale_loss_no_uni(self):
        # Sale loss with AAB == original basis → entire loss is capital
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                    "Date: Sale": _dt(2023, 6, 15),
                    "Price per share: Sale": 40.0,
                    "Exchange Rate: Sale": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 60.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.sale_gain_loss == -1000
        assert result.adjusted_basis == 5000
        assert result.capital_loss == -1000


class TestSaleInPriorYear:
    def test_prior_year_sale_skipped(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2021),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                    "Date: Sale": _dt(2022, 3, 15),
                    "Price per share: Sale": 60.0,
                    "Exchange Rate: Sale": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2021, 55.0, 1.0), (2022, 60.0, 1.0), (2023, 65.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result.skipped is True


# -- Part I --


class TestPart1:
    def test_single_lot_value(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 60.0, 1.0)])
        result = compute_part1(df_lot, df_eoy, 2023, FILEPATH)
        assert result.unsold_shares == 100
        assert result.value_of_pfic == 6000

    def test_multiple_lots_value(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                },
                {
                    "Date: Acquisition": _dt(2023, 6, 1),
                    "Price per share: Acquisition": 55.0,
                    "Number of shares": 200,
                    "Cost: Acquisition": 11000.0,
                    "Exchange Rate: Acquisition": 1.0,
                },
            ]
        )
        df_eoy = _make_eoy_df([(2023, 60.0, 1.0)])
        result = compute_part1(df_lot, df_eoy, 2023, FILEPATH)
        assert result.unsold_shares == 300
        assert result.value_of_pfic == 18000


# -- Prior year losses force AAB != FMV --


class TestDeniedLossPreservesBasis:
    # Denied loss (no UNI) leaves AAB unchanged. Next year's AAB(BOY) must
    # reflect that, not use prior-year FMV.
    # 2023: acquire 100@$100, FX1.0 → loss $2000 denied (UNI=0) → AAB stays $10000
    # 2024: price $110, FMV=$11000 → gain = $11000-$10000 = $1000
    def test_denied_loss_preserves_basis(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 100.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 10000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 80.0, 1.0), (2024, 110.0, 1.0)])

        result_2023 = compute_lot(df_lot, df_eoy, 0, 2023, FILEPATH)
        assert result_2023.gain_loss == -2000
        assert result_2023.adjusted_basis == 10000
        assert result_2023.unreversed == 0
        assert result_2023.ordinary_loss == 0

        result_2024 = compute_lot(df_lot, df_eoy, 0, 2024, FILEPATH)
        assert result_2024.adjusted_basis == 10000
        assert result_2024.gain_loss == 1000


class TestPartialLossFollowedByGain:
    # Loss partially absorbed by UNI, then a gain year.
    # 2022: gain $1000 → AAB=$6000, UNI=$1000
    # 2023: loss $1500, allowed $1000 → AAB=$5000, UNI=$0
    # 2024: FMV=$5500 → gain = $5500-$5000 = $500
    def test_aab_after_partial_loss(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2022),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2022, 60.0, 1.0), (2023, 45.0, 1.0), (2024, 55.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2024, FILEPATH)
        assert result.adjusted_basis == 5000
        assert result.gain_loss == 500


class TestLossMisclassifiedAfterPriorLoss:
    # After partial loss absorption, a small FMV increase still represents a loss.
    # 2022: gain $1000 → AAB=$6000, UNI=$1000
    # 2023: loss $1500, allowed $1000 → AAB=$5000, UNI=$0
    # 2024: FMV=$4700 → AAB=$5000 → non-deductible loss of $300 (not a gain)
    def test_loss_not_gain_after_prior_loss(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2022),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2022, 60.0, 1.0), (2023, 45.0, 1.0), (2024, 47.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2024, FILEPATH)
        assert result.adjusted_basis == 5000
        assert result.gain_loss == -300
        assert result.unreversed == 0
        assert result.ordinary_loss == 0


class TestMultiYearCarryforward:
    # Figure 6 from the PFIC.xyz article:
    # 2019: gain $2000 → AAB=$7000, UNI=$2000
    # 2020: loss $1500, allowed $1500 → AAB=$5500, UNI=$500
    # 2021: loss $1000, allowed $500 → AAB=$5000, UNI=$0
    # 2022: gain $1500 → AAB=$6500, UNI=$1500
    def test_2020_loss_after_gain(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2019),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2019, 70.0, 1.0), (2020, 55.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2020, FILEPATH)
        assert result.adjusted_basis == 7000
        assert result.fmv == 5500
        assert result.gain_loss == -1500
        assert result.unreversed == 2000
        assert result.ordinary_loss == -1500

    def test_2021_loss_after_partial_absorption(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2019),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2019, 70.0, 1.0), (2020, 55.0, 1.0), (2021, 45.0, 1.0)])
        result = compute_lot(df_lot, df_eoy, 0, 2021, FILEPATH)
        assert result.adjusted_basis == 5500
        assert result.fmv == 4500
        assert result.gain_loss == -1000
        assert result.unreversed == 500
        assert result.ordinary_loss == -500

    def test_2022_gain_after_losses(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2019),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df(
            [
                (2019, 70.0, 1.0),
                (2020, 55.0, 1.0),
                (2021, 45.0, 1.0),
                (2022, 65.0, 1.0),
            ]
        )
        result = compute_lot(df_lot, df_eoy, 0, 2022, FILEPATH)
        assert result.adjusted_basis == 5000
        assert result.gain_loss == 1500


class TestFXRateAfterDeniedLoss:
    # FX change after a denied loss.
    # 2023: acquire 100@€50, FX1.0 → basis $5000; price €40, FX1.0 → denied loss
    # 2024: price €55, FX1.2 → FMV=$6600, AAB=$5000 → gain $1600
    def test_fx_with_denied_loss(self):
        df_lot = _make_lot_df(
            [
                {
                    "Date: Acquisition": _dt(2023),
                    "Price per share: Acquisition": 50.0,
                    "Number of shares": 100,
                    "Cost: Acquisition": 5000.0,
                    "Exchange Rate: Acquisition": 1.0,
                }
            ]
        )
        df_eoy = _make_eoy_df([(2023, 40.0, 1.0), (2024, 55.0, 1.2)])
        result = compute_lot(df_lot, df_eoy, 0, 2024, FILEPATH)
        assert result.adjusted_basis == 5000
        assert result.gain_loss == 1600


# -- LotResult.lot_summary aggregation --


class TestLotSummaryAggregation:
    def test_holding_gain_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=True,
            fmv=6000,
            adjusted_basis=5000,
            original_basis=5000.0,
            gain_loss=1000,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 1000,
            "ordinary_losses": 0,
            "capital_losses": 0,
        }

    def test_holding_ordinary_loss_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=True,
            fmv=4000,
            adjusted_basis=7000,
            original_basis=5000.0,
            gain_loss=-3000,
            unreversed=2000,
            ordinary_loss=-2000,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 0,
            "ordinary_losses": 2000,
            "capital_losses": 0,
        }

    def test_holding_denied_loss_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=True,
            fmv=4000,
            adjusted_basis=5000,
            original_basis=5000.0,
            gain_loss=-1000,
            unreversed=0,
            ordinary_loss=0,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 0,
            "ordinary_losses": 0,
            "capital_losses": 0,
        }

    def test_sale_gain_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=False,
            fmv=0,
            adjusted_basis=5000,
            original_basis=5000.0,
            gain_loss=0,
            proceeds=6500,
            sale_gain_loss=1500,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 1500,
            "ordinary_losses": 0,
            "capital_losses": 0,
        }

    def test_sale_ordinary_loss_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=False,
            fmv=0,
            adjusted_basis=7000,
            original_basis=5000.0,
            gain_loss=0,
            proceeds=5500,
            sale_gain_loss=-1500,
            unreversed=2000,
            ordinary_loss=-1500,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 0,
            "ordinary_losses": 1500,
            "capital_losses": 0,
        }

    def test_sale_capital_loss_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=False,
            fmv=0,
            adjusted_basis=5000,
            original_basis=5000.0,
            gain_loss=0,
            proceeds=4000,
            sale_gain_loss=-1000,
            unreversed=None,
            ordinary_loss=0,
            capital_loss=-1000,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 0,
            "ordinary_losses": 0,
            "capital_losses": 1000,
        }

    def test_skipped_summary(self):
        lr = LotResult(
            lot_index=0,
            is_holding=False,
            fmv=0,
            adjusted_basis=0,
            original_basis=0,
            gain_loss=0,
            skipped=True,
        )
        assert lr.lot_summary == {
            "ordinary_gains": 0,
            "ordinary_losses": 0,
            "capital_losses": 0,
        }
