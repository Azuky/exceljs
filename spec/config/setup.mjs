import { createRequire } from 'node:module';
import { chai } from 'vitest';
import chaiDatetime from 'chai-datetime';
import chaiXmlOrdered from '../utils/chai-xml-ordered.js';

globalThis.verquire = createRequire(import.meta.url)('../utils/verquire');

chai.use(chaiXmlOrdered);
chai.use(chaiDatetime);
