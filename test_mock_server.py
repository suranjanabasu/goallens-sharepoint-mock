import json
import threading
import unittest
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from mock_server import DRIVE_ID, SITE_ID, create_server


class GraphMockTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = create_server(port=0, quiet=True)
        cls.base = f"http://127.0.0.1:{cls.server.server_port}"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=2)

    def setUp(self):
        self.post("/__admin/reset", {})

    def request(self, path, token=True):
        headers = {"Authorization": "Bearer test"} if token else {}
        return urlopen(Request(self.base + path, headers=headers), timeout=2)

    def get_json(self, path, token=True):
        with self.request(path, token=token) as response:
            return json.load(response)

    def post(self, path, body):
        request = Request(
            self.base + path,
            data=json.dumps(body).encode(),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urlopen(request, timeout=2) as response:
            return json.load(response)

    def test_token_endpoint(self):
        result = self.post("/mocktenant/oauth2/v2.0/token", {})
        self.assertEqual(result["access_token"], "mock-sharepoint-access-token")

    def test_site_and_drives(self):
        site = self.get_json("/v1.0/sites/goallensdev.sharepoint.com:/sites/GoalLensSandbox")
        self.assertEqual(site["id"], SITE_ID)
        drives = self.get_json(f"/v1.0/sites/{SITE_ID}/drives")
        self.assertEqual(drives["value"][0]["id"], DRIVE_ID)

    def test_delta_pagination_and_cursor(self):
        first = self.get_json(f"/v1.0/drives/{DRIVE_ID}/root/delta")
        self.assertIn("@odata.nextLink", first)
        second_url = first["@odata.nextLink"].replace(self.base, "")
        second = self.get_json(second_url)
        self.assertIn("@odata.deltaLink", second)
        self.assertEqual(len(first["value"]) + len(second["value"]), 2)

    def test_new_version_scenario_and_download(self):
        self.post("/__admin/scenario", {"scenario": "new_version"})
        delta = self.get_json(f"/v1.0/drives/{DRIVE_ID}/root/delta?token=baseline")
        self.assertEqual(delta["value"][0]["eTag"], '"etag-scorecard-v2"')
        with self.request(f"/v1.0/drives/{DRIVE_ID}/items/file-scorecard/content") as response:
            self.assertIn(b"Demo-to-Close Conversion", response.read())

    def test_throttles_once(self):
        self.post("/__admin/scenario", {"scenario": "throttled_once"})
        path = f"/v1.0/drives/{DRIVE_ID}/root/delta"
        with self.assertRaises(HTTPError) as error:
            self.request(path)
        self.assertEqual(error.exception.code, 429)
        self.assertEqual(error.exception.headers["Retry-After"], "1")
        error.exception.close()
        self.assertIn("value", self.get_json(path))

    def test_permission_denied(self):
        self.post("/__admin/scenario", {"scenario": "permission_denied"})
        with self.assertRaises(HTTPError) as error:
            self.request(f"/v1.0/drives/{DRIVE_ID}/root/delta")
        self.assertEqual(error.exception.code, 403)
        error.exception.close()

    def test_expired_delta_requires_resync(self):
        self.post("/__admin/scenario", {"scenario": "expired_delta"})
        with self.assertRaises(HTTPError) as error:
            self.request(f"/v1.0/drives/{DRIVE_ID}/root/delta?token=baseline")
        self.assertEqual(error.exception.code, 410)
        error.exception.close()

    def test_delete_and_restore_delta_shapes(self):
        self.post("/__admin/scenario", {"scenario": "deleted"})
        deleted = self.get_json(f"/v1.0/drives/{DRIVE_ID}/root/delta?token=baseline")
        self.assertEqual(deleted["value"][0]["deleted"]["state"], "deleted")

        self.post("/__admin/scenario", {"scenario": "restored"})
        restored = self.get_json(f"/v1.0/drives/{DRIVE_ID}/root/delta?token=after-delete")
        self.assertEqual(restored["value"][0]["eTag"], '"etag-scorecard-v1"')

    def test_download_failure(self):
        self.post("/__admin/scenario", {"scenario": "download_failure"})
        with self.assertRaises(HTTPError) as error:
            self.request(f"/v1.0/drives/{DRIVE_ID}/items/file-scorecard/content")
        self.assertEqual(error.exception.code, 500)
        error.exception.close()


if __name__ == "__main__":
    unittest.main()
