<?php
/**
 * Low Stock -- a storefront module written from nothing.
 *
 * Chapter 19. This is the smallest honest example of extending the platform:
 * it adds something the shop could not do before, it does not edit a single
 * file that OpenCart ships, and it survives an upgrade.
 *
 * What it does: lists enabled products whose stock has fallen below a
 * threshold, as a "last few left" panel. Useful in itself, and chosen for
 * this course because its cost is exactly one database query -- which makes
 * Chapter 19's measurement of what a module costs a page unambiguous.
 *
 * Three things to notice in the five lines below the comment:
 *
 *   1. THE NAMESPACE IS THE PATH. Opencart\Catalog\Controller\Extension\
 *      Ecadmod\Module maps to extension/ecadmod/catalog/controller/module/.
 *      Get it wrong and the loader reports that the class does not exist,
 *      which is true and unhelpful.
 *
 *   2. index() TAKES ITS SETTINGS AS AN ARGUMENT. The layout decides which
 *      instance of the module is being rendered and hands it its own row of
 *      settings. A module that reads $this->config instead is a module that
 *      cannot be placed twice with different settings.
 *
 *   3. IT RETURNS A STRING. It does not print, and it does not set a
 *      response. The layout decides where the string goes.
 */
namespace Opencart\Catalog\Controller\Extension\Ecadmod\Module;

class LowStock extends \Opencart\System\Engine\Controller {
	/**
	 * @param array<string, mixed> $setting this instance's own settings
	 *
	 * @return string the module's HTML, or '' to render nothing
	 */
	public function index(array $setting): string {
		$this->load->language('extension/ecadmod/module/lowstock');

		$data['heading_title'] = $this->language->get('heading_title');
		$data['text_left'] = $this->language->get('text_left');

		$threshold = (int)($setting['threshold'] ?? 5);
		$limit     = (int)($setting['limit'] ?? 4);

		// One query, and it is written here rather than in a model only
		// because the whole point of the chapter is that you can see what the
		// module costs. A real module of any size puts this in a model.
		$query = $this->db->query("SELECT p.product_id, p.quantity, pd.name FROM `" . DB_PREFIX . "product` p LEFT JOIN `" . DB_PREFIX . "product_description` pd ON (p.product_id = pd.product_id) WHERE p.status = '1' AND p.quantity > '0' AND p.quantity <= '" . (int)$threshold . "' AND pd.language_id = '" . (int)$this->config->get('config_language_id') . "' ORDER BY p.quantity ASC LIMIT " . (int)$limit);

		$data['products'] = [];

		foreach ($query->rows as $row) {
			$data['products'][] = [
				'name'     => $row['name'],
				'quantity' => $row['quantity'],
				// Build links with the url object, never by string
				// concatenation: it carries the language and the session
				// through, and Chapter 21 depends on that.
				'href'     => $this->url->link('product/product', 'language=' . $this->config->get('config_language') . '&product_id=' . $row['product_id'])
			];
		}

		if (!$data['products']) {
			return '';
		}

		return $this->load->view('extension/ecadmod/module/lowstock', $data);
	}
}
