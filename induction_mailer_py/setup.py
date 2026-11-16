from setuptools import setup


if __name__ == "__main__":
    # Legacy setuptools entrypoint.
    # We intentionally avoid pyproject.toml / build isolation so installs work
    # in offline / proxy-restricted environments (common on corporate Windows).
    setup()
