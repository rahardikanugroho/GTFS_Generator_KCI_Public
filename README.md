# GTFS_Generator_Public

A small, dependency-free Python tool that generates a valid [GTFS](https://gtfs.org/) (General Transit Feed Specification) feed using synthetic data. Useful for learning, testing and demonstrating GTFS-based transit tools without relying on real-world data.

## What it generates

| File | Content |
|------|---------|
| `agency.txt` | Single demo agency |
| `stops.txt` | Stops placed along each route |
| `routes.txt` | Route definitions (bus type) |
| `trips.txt` | Trip per route per day |
| `stop_times.txt` | Ordered arrival/departure times per stop |
| `calendar.txt` | Weekday service definition |
| `feed_info.txt` | Feed metadata |

## Quick start

```bash
python generate_gtfs.py --out output --zip
```

### CLI options

| Flag | Default | Description |
|------|---------|-------------|
| `--out` | `output` | Output folder |
| `--zip` | `False` | Also create a `.zip` of the feed |
| `--routes` | `3` | Number of routes |
| `--stops-per-route` | `6` | Stops per route |
| `--trips-per-day` | `5` | Departures per day per route |
| `--seed` | `42` | Random seed for reproducibility |
| `--start-date` | auto | Feed start date (YYYY-MM-DD) |
| `--config` | — | Optional JSON config file |

## Example

```bash
python generate_gtfs.py --routes 5 --stops-per-route 8 --trips-per-day 10 --zip
```

Generates 5 routes, 40 stops, 50 daily trips and writes the feed to `output/` + `output.zip`.

## Use in larger projects

Pair with [GTFS_Monitoring_Public](https://github.com/rahardikanugroho/GTFS_Monitoring_Public) to validate the generated feed.

## License

MIT