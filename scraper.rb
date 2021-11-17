# frozen_string_literal: true

require 'nokogiri'
require 'rubyXL'
require 'down'
require 'fileutils'
require 'watir'
require 'net/http'
require 'uri'
require 'json'
require 'byebug'

IMGUR_CLIENT_ID='2c8bcbf4afd0d38'
IMGUR_UPLOAD_URL='https://api.imgur.com/3/upload'
@new_row_index = 3
@page = 93
@per_page = 10

def current_spreadsheet
  @current_spreadsheet ||= RubyXL::Parser.parse('./tambah-sekaligus(1366122).xls')
end

def category_sheet
  @category_sheet ||= begin
    result = {}
    list = current_spreadsheet.worksheets[1]
    list.each_with_index do |row, index|
      key = row[2].value
      val = row[3].value
      result[key] = val.to_i
    end
    result
  end
end

def content_sheet
  current_spreadsheet.worksheets[0]
end

# def img_download(img_url)
#   tempfile = Down.download(img_url)
#   img_path = "./images/product_name.jpg"
#   FileUtils.mv(tempfile.path, img_path)
# end

def upload_imgur(url, title)
  uri = URI.parse(IMGUR_UPLOAD_URL)
  # 'Authorization': "Client-ID #{IMGUR_CLIENT_ID}"
  headers = {
    'Authorization': "Client-ID #{IMGUR_CLIENT_ID}",
    'Content-Type' =>'application/json',
    'Accept'=>'application/json'
  }

  http = Net::HTTP.new(uri.host, uri.port)
  http.use_ssl = uri.port == 443
  request = Net::HTTP.post_form(uri, image: url, type: 'url', title: title)
  response = JSON.parse(request.body, object_class: OpenStruct)
  response.data.link
end

def scrap_page(url)
  browser = Watir::Browser.new
  browser.goto url

  # scroll down & wait 2 seconds
  browser.scroll.to :bottom
  sleep(2)

  result = Nokogiri::HTML.parse(browser.html)
  browser.close

  result
end

def image_big(url, title)
  img_url = url.gsub('100-square', '500-square')
  img_url = img_url.split('.webp')
  cached_img = img_url.first

  begin
    upload_imgur(cached_img, title)
  rescue
    cached_img
  end
end

def generate_desc(original_desc)
  additional = "<br><br>Semua barang ready stock yah.<br>
  Pemesanan hari senin - sabtu sebelum pukul 16.00 akan kami kirim di hari yang sama dan diatas itu dikirim keesokan harinya.<br>
  Pemesanan hari Minggu dan tanggal merah akan diproses pada hari kerja berikutnya.<br>
  <br>
  Kami menjual produk spare part maupun accecories serta kebutuhan untuk offroad khusus Jeep.<br>
  Silahkan chat kami bila ada pertanyaaan mengenai produk yang sudah ada di list produk maupun yang belum ada di list produk kami.<br>
  <br>
  Selamat berbelanja di toko kami 🙏"

  "#{original_desc}#{additional}".dup.gsub!("<br>", "\x0A")
end

def price_with_margin(price)
  percentage = 50 if price <= 50_000
  percentage = 30 if price > 50_000 && price <= 100_000
  percentage = 20 if price > 100_000 && price <= 1_000_000
  percentage = 10 if price > 1_000_000 && price <= 2_500_000
  percentage = 7.5 if price > 2_500_000 && price <= 5_000_000
  percentage = 5 if price > 5_000_000
  percentage_val = percentage.to_f / 100
  price.to_f + (percentage_val * price.to_f)
end

def fetch_page(page)
  latest_row = @new_row_index
  # variables
  name = page.css('.css-v7vvdw').text
  name = capitalize(name)
  desc = generate_desc(page.css('.css-168ydy0 div').text)
  desc_detail = page.css('.css-1vbldqk')
  cat_text = desc_detail[2].css('a').text
  cat_id = category_sheet[cat_text]
  weight_text = desc_detail[1].css('span.main').text
  weight = weight_text.tr('^0-9', '').to_i
  condition = desc_detail[0].css('span.main').text
  price_text = page.css('.css-32gaxy .price').text
  price = price_text.tr('^0-9', '').to_i
  selling_price = price_with_margin(price).to_i
  stock = page.css('.css-9vy0ue p b').text
  stock = 1 if stock == '999,9rb'
  active = stock.to_i.zero? ? 'Non Aktif' : 'Aktif'
  etalase_no = 27355886
  images = page.css('.css-ikbtkl').each_with_index.map{|container, idx| image_big(container.css('img').first['src'], parameterize(name)) }
  min_order = 1
  asuransi = 'opsional'

  # print results
  puts "Nama : #{name}"
  puts "Deskripsi : #{desc}"
  puts "Kategori Kode : #{cat_id}"
  puts "Berat (gr) : #{weight}"
  puts "Minimum Order : #{min_order}"
  puts "Nomor Etalase : #{etalase_no}"
  puts "Kondisi : #{condition}"
  puts "Gambar (X): #{images}"
  puts "Status : #{active}"
  puts "Jumlah Stok : #{stock}"
  puts "Harga : #{selling_price}"
  puts "Asuransi Pengiriman : #{asuransi}"

  # save results to excel
  new_row = [nil, name, desc, cat_id, weight, min_order, etalase_no, nil, condition, images[0], images[1], images[2], images[3], images[4], nil, nil, nil, nil, active, stock, selling_price, asuransi]
  new_row.each_with_index do |value, index|
    content_sheet.add_cell(latest_row, index, value)
  end
  @new_row_index += 1
end

def perform
  total_scrapped = 0
  # total_products = 1134
  per_page = @per_page
  # total_scraps = (total_products.to_f/per_page.to_f).ceil
  # total_scraps = 2

  # loop every page
  # total_scraps.times do |index|
  #   page = index + 1
    page = @page
    puts '=' * 10
    puts "Page #{page}"
    puts '=' * 10
    url = "https://www.tokopedia.com/parts-jeep/page/#{page}?perpage=#{per_page}"
    pagination_parsed_page = scrap_page(url)
    products = pagination_parsed_page.css('.css-1sn1xa2')
    # loop every product in one page
    products.each_with_index do |product, product_idx|
      next if product_idx > 2
      product_url = product.css('.css-1ehqh5q a').first['href']
      puts "Product URL: #{product_url}"
      parsed_page = scrap_page(product_url)
      fetch_page(parsed_page)
      total_scrapped += 1
      puts '===================='
      # break
    end
  #   break
  # end

  current_spreadsheet.write("tokped_#{Time.now.strftime('%y%m%d')}.xlsx")
  puts "Total Scrapped : #{total_scrapped} products"
end

def parameterize(text_string)
  text_string.downcase.gsub(' ', '-')
end

def capitalize(text_string)
  text_string.split(' ').map{|t| t.length > 2 ? t.capitalize() : t.upcase }.join(' ')
end

perform
