# Exchange Web Services Cleaner

A powerful utility for cleaning and managing Exchange mailbox folders via Exchange Web Services (EWS) with real-time monitoring of EWS operations.

![EWS Cleaner](https://i.imgur.com/tITu8hD.png)

## Features

- Fast cleaning of Exchange mailbox folders
- Real-time monitoring of EWS API calls
- Performance statistics for EWS operations
- User-friendly interface with colored output
- Error handling with intelligent retry mechanisms
- Impersonation support for administrator accounts
- Cross-platform compatibility (Windows and Linux)

## Requirements

- Python 3.6 or higher
- Exchange server with EWS access enabled
- Network access to your Exchange server
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone or download the repository:

```bash
git clone https://github.com/yourusername/exchange-cleaner.git
```

2. Install the required dependencies:

```bash
pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ ews-folder-cleaner==1.0.0
```

## Usage

### Basic Usage

Run the script without parameters to be prompted for credentials:

```bash
python exchange_cleaner_linux_fixed.py
```

The script will prompt you for:
- Your email address
- Your password
- Whether to impersonate another user (optional)

### Command Line Arguments

You can also provide credentials and options via command line:

```bash
python exchange_cleaner_linux_fixed.py [OPTIONS] [username] [password] [impersonated_user]
```

### Available Options

| Option | Description |
|--------|-------------|
| `--server SERVER` | Specify the Exchange server URL (required) |
| `--no-log-window` | Disable separate log window |
| `--console-log` | Display logs in main console |
| `--no-stats-window` | Disable statistics window |
| `--classic-ui` | Use classic interface instead of Rich unified interface |
| `--auto-monitor` | Automatically start EWS monitoring |
| `-h, --help` | Display help information |

## Interface Commands

While running, the following commands are available:

- `1-9`: Select a folder by its number
- `m`: Enable/disable EWS monitoring
- `s`: Display EWS statistics
- `q`: Exit the program
- `h` or `?`: Display help

## Monitoring Interface

The monitoring interface displays:

1. **Current Processing Progress**:
   - Folder name
   - Items processed
   - Items remaining
   - Processing speed
   - Estimated time remaining

2. **Progress History**:
   - Historical data of folder processing

3. **EWS Statistics**:
   - Active calls
   - Total calls
   - Min/Avg/Max times for API calls

4. **Recent Logs**:
   - Last 5 log entries with severity indicators

5. **EWS Request Monitoring**:
   - Recent API calls with timing information

## Error Handling

The script implements intelligent error handling:

- **Server Busy**: Automatically waits with appropriate back-off times
- **Database Unavailable**: Waits for 30 seconds before retrying
- **Slow Calls**: Adds delays for operations taking more than 1000ms
- **Network Issues**: Graceful handling with appropriate retries

## Advanced Features

### Impersonation

For admin accounts, you can impersonate another user to manage their mailbox:

```bash
python exchange_cleaner_linux_fixed.py username password impersonated_user@domain.com
```

### Custom Server

Specify a custom Exchange server:

```bash
python exchange_cleaner_linux_fixed.py --server outlook.office365.com username password
```

## Troubleshooting

### Common Issues

1. **Connection Failures**:
   - Verify your network connectivity
   - Check that your Exchange server is accessible
   - Confirm EWS is enabled on your server

2. **Authentication Errors**:
   - Verify your username and password
   - Check if your account requires MFA (not supported directly)
   - Try using an app password if MFA is enabled

3. **Performance Issues**:
   - Large folders may take time to process
   - The script implements throttling to avoid server overload
   - Use the monitoring interface to track progress

### Log Files

Log files are stored in:
- Windows: `%TEMP%\ews_logs.txt`
- Linux: `/tmp/ews_logs.txt`

View logs in real-time using:

```bash
tail -f /tmp/ews_logs.txt  # Linux
```

## License

This software is released under the MIT License.

## Credits

Developed by [Your Name/Organization]

Using libraries:
- exchangelib
- rich
- colorama 