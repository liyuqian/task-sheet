import { archive, createSheet, filterDueSoon, onEdit, onOpen } from './sheet';
import { testAll } from '../test/integration_test';

function main() {
  createSheet();
}

// Functions archive, filterDueSoon, testAll, onOpen and onEdit are needed but
// not explicitly called inside main. So we create this function to introduce
// the dependencies. This function is not expected to be called anywhere.
function _deps() {
  onOpen();
  onEdit(null);
  testAll();
  filterDueSoon(null);
  archive(null);
}
